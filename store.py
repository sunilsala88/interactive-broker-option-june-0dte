

#write a python code to store data in pickle file
import pickle


def store(data):
    pickle.dump(data,open('data.pickle','wb'))

def load():
    return pickle.load(open('data.pickle', 'rb'))

try:
    data=load()
    print(data)
    print('data loaded from pickle file')
except:
    print('no data to read')

    
