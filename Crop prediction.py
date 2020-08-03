
# coding: utf-8

# In[1]:


import numpy as np


# In[2]:


import pandas as pd
import matplotlib.pyplot as plt
get_ipython().run_line_magic('matplotlib', 'inline')
import seaborn as sns


# In[3]:


df = pd.read_csv('D:/crop.csv')
df


# In[4]:


df.describe()


# In[5]:


df.info()


# In[6]:


plt.figure(figsize=(18,6))
sns.distplot(df['Avg Temperature(celcius)'])


# In[7]:


sns.pairplot(df)


# In[8]:


from sklearn.cross_validation import train_test_split


# In[10]:


X = df.drop(['Weather','crop_cultivated'],axis=1,inplace=True)
y = df['crop_cultivated']


# In[ ]:


X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.29)


# In[ ]:


from sklearn.tree import DecisionTreeClassifier


# In[ ]:


dtree = DecisionTreeClassifier()


# In[ ]:


dtree.fit(X_train,y_train)

