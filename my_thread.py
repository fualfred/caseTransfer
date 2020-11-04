# -*- coding: utf-8 -*-
import threading
class My_thread(threading.Thread):

    def __init__(self, func, args):
        super(My_thread,self).__init__()
        self.func = func
        self.args = args
        print(self.func)
        print(self.args)
    def run(self):
        self.result = self.func(self.args)
    
    def get_result(self):
        try:
            return self.result
        except Exception:
            return None