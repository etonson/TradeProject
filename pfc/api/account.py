'''
Created on 2022年7月11日

@author: eton
'''

import comtypes.client


class Account:

    def __init__(self, account, password, ip_address):
        self.client = comtypes.client.CreateObject("COM_PFCFAPI.COM_PFCFAPI")
        self.account = account
        self.password = password
        self.ip_address = ip_address
    def login(self):
        self.client.PFCLogin(self.account, self.password, self.ip_address)
        ACTNO =self.client.UserOrderSet[0]
        return ACTNO
    def logout(self):
        self.client.PFCLogout()
    
