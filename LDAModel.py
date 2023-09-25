import nltk
# # nltk.download('stopwords')
# # nltk.download('punkt')
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
stop_words = set(stopwords.words('english'))
stop_words.update(['general','procedures','products','execution','section','sections','mm','one','includes','summary'])

from gensim.models import LdaModel
from gensim import corpora, models
# import gensim
import sqlite3
import pickle
# import pyLDAvis.gensim_models as gensimvis
# import pyLDAvis
# import numpy as np
import time
from datetime import date

def preprocess(text):
    tokens=[]
    filtered_tokens=[]
    master=[]
    for i in range(len(text)):
        tokens=word_tokenize(text[i][1].lower())
        # filtered_tokens = [token for token in tokens if token.isalpha() and token not in stop_words]
        for token in tokens:
            if token.isalpha() and token not in stop_words:
                    filtered_tokens.append(token)
        tokens=[]
        master.append(filtered_tokens)
        filtered_tokens=[]
    
    bigram_phrases=models.Phrases(master,min_count=5,threshold=100)
    trigram_phrases=models.Phrases(bigram_phrases[master],threshold=100)

    bigram=models.phrases.Phraser(bigram_phrases)
    trigram=models.phrases.Phraser(trigram_phrases)
    
    # Apply the models to the data
    data_bigrams = [bigram[doc] for doc in master]
    data_trigrams = [trigram[bigram[doc]] for doc in master]

    return data_trigrams

def est_access():
    DB_Path=r"C:\Users\logan\Desktop\core\Components.db"
    global conn
    global cur
    conn = sqlite3.connect(DB_Path) 
    cur=conn.cursor()

 
def start_training():
    #checks if we need to remake the table with a new column? (ie first time training)
    cur.execute("PRAGMA table_info(LDABlobs)")
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='LDABlobs'")
    exists = cur.fetchall()
    if len(exists)==0: 
        buildTable(cur,conn)
    
    cur.execute('SELECT Subdivision_Description, GROUP_CONCAT(Details) AS all_Details FROM Components GROUP BY Subdivision_Description ORDER BY ID')
    results = cur.fetchall()

    print('done adding details to trainable list')
    return results

def filtered2LDA(documents):
    dictionary = corpora.Dictionary(documents)
    dictionary.filter_extremes(no_below=0, no_above=0.35)
    dictionary.compactify()
    BOWcorpus = [dictionary.doc2bow(doc) for doc in documents]
    ##train LDA model
    lda_model = LdaModel(corpus=BOWcorpus, num_topics=50, id2word=dictionary, passes=10,random_state=100,update_every=1,chunksize=100,alpha="auto")
    
    lda_blob=pickle.dumps(lda_model)
    cur.execute("SELECT COUNT(*) FROM LDABlobs")
    row = cur.fetchone()[0]
    cdate=date.today()
    date_string = cdate.strftime('%Y-%m-%d')
    # cur.execute("INSERT INTO LDABlobs (Date, LDA_blob) VALUES('{date_string}', {lda_blob})")
    cur.execute("INSERT INTO LDABlobs (Date, LDA_blob) VALUES (?, ?)", (date_string, lda_blob))
    conn.commit()
    
    print('done training LDA model - saved to db')
        

def buildTable(cur,conn):
    print('No LDA_blob table -- building now')
    try: 
        cur.execute(''' CREATE TABLE LDABlobs (
        Date TEXT,
        LDA_blob BLOB)
        ''')
        
    except: 
        print("Something didnt work when building new Components table.")

def train():
    start=time.time()
    data=start_training()
    tokens=preprocess(data) #filters and tokenizes ALL 'docs' (db entries), in one master, list of lists
    filtered2LDA(tokens) #takes filtered tokens, 
    end=time.time()
    print('It took: ',end-start,' seconds to train')
    

def retrieve_fromDB():
    cur.execute("SELECT COUNT(*) FROM LDABlobs")
    endLDACol = cur.fetchone()[0]
    
    MAX_num_associated=5
    
    cur.execute('''SELECT LDA_blob FROM LDABlobs ORDER BY ROWID DESC LIMIT 1''')
    tmp=cur.fetchone()[0]
    preTrained_lda_model=pickle.loads(tmp)
    
    return preTrained_lda_model

def test_performance(lda_model,inputWords,n=3,thresh=0.5):
    preprocessed_inputwords = preprocess(inputWords)
    original_dictionary=lda_model.id2word
    # Convert the preprocessed words to a bag-of-words representation using the original dictionary used to train the model
    input_corpus = [original_dictionary.doc2bow(word) for word in preprocessed_inputwords]

    #get original corpus the lda model was trained on
    cur.execute('SELECT Subdivision_Description, GROUP_CONCAT(Details) AS all_Details FROM Components GROUP BY Subdivision_Description ORDER BY ID')
    ckdata = cur.fetchall()
    processed_ckdata=preprocess(ckdata)
    
    Intelligence=[]
    # Get the topic distribution for the input words
    
    # for i in input_corpus:
    for i in range(0,len(input_corpus)):
        if len(input_corpus[i])==0:
            continue
        print(input_corpus[i][0])
        topic_distribution = lda_model.get_document_topics(input_corpus[i]) 
        # Sort the topic distribution by probability
        sorted_topics = sorted(topic_distribution, key=lambda x: x[1], reverse=True) 
        topicID=sorted_topics[0][0]
        ######Now find a doc
 
        related_documents = []
        for document in processed_ckdata:
            if document == []:
                continue 
            document_vector = lda_model.id2word.doc2bow(document)
            topic_distribution = lda_model[document_vector]
            
            for topic, prob in topic_distribution:
                if topic == topicID: #must be more than 50% related
                    related_documents.append([prob, document])  
                    
        orderedSortedDocs=sorted(related_documents,key=lambda x: x[0], reverse=True)
        
        Intelligence.append([inputWords[i],orderedSortedDocs])
    
    # n=3 #get top n documents related to each keyword
    # thresh=0.5 #pct related threshold
    
    findIndex=[]
    for result in Intelligence: 
        docs=result[1]
        if len(docs)<n:
            n=len(docs)
            
        for v in range(0,n):
            if docs[v][0]>thresh:
                 findIndex.append(docs[v][1])

    ids=[] ##THESE ARE DOCUMENT IDS NOT INDECIES FOR DB              
    for green in findIndex:
        id=processed_ckdata.index(green)
        ids.append(id)       
    
    SubsectionNames=[]
    for val in ids:
        SubsectionNames.append(ckdata[val][0])
    
    return SubsectionNames

                    
# est_access()
# # train() #this does the training
# TrainedLDAmodel=retrieve_fromDB()
# inputwords=['hydronic','hydronic','plumbing','ventilation']
# Ids=test_performance(TrainedLDAmodel,inputwords)
# print(Ids)