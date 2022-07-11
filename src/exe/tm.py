'''
Created on 2022年7月10日

@author: eton
'''
from pfc.api.account import Account

user = Account("account","your password","test225.pfctrade.com")
a=user.login()
print(a)