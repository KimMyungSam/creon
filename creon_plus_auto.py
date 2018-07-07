
# coding: utf-8

# In[3]:


from pywinauto import application
from pywinauto import timings
import time
import os

app = application.Application()
#app.start('C:\CREON\STARTER\coStarter.exe /prj:cp /id:creo03 /pwd:rlaehgus /pwdcert:zzzzz /autostart')
app.start('C:\CREON\STARTER\coStarter.exe /prj:cp /id:creo03 /pwd:rlaehgus /autostart')

time.sleep(50)

os.system("taskkill /im coStarter.exe")

