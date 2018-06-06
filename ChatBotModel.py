
# coding: utf-8

# In[2]:


import telegram
from telegram.ext import Updater, CommandHandler

class TelegramBot:
    def __init__(self, name, token):
        self.core = telegram.Bot(token)
        self.updater = Updater(token)
        self.id = 571675744
        self.name = name

    def sendMessage(self, text):
        self.core.sendMessage(chat_id = self.id, text=text)
        
    def sendPhoto(self, photo):
        self.core.sendPhoto(chat_id = self.id, photo=photo)

    def stop(self):
        self.updater.start_polling()
        self.updater.dispatcher.stop()
        self.updater.job_queue.stop()
        self.updater.stop()


# In[5]:


class Bot2ndBUS (TelegramBot):
    def __init__(self):
        self.token = '472813594:AAGpKj5sn4gkATwB19oHHnsjcKmZb0EJ5S4'
        TelegramBot.__init__(self, '2ndBUS', self.token)
        self.updater.stop()

    def add_handler(self, cmd, func):
        self.updater.dispatcher.add_handler(CommandHandler(cmd, func))

    def start(self):
        self.sendMessage('2ndBUS RUN.')
        self.updater.start_polling()
        self.updater.idle()

