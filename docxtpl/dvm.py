# -*- coding: utf-8 -*-
"""
Created: 2022-11-3
@author: tsy
"""

class DVM:
    """ load list of block tuple, list of fill color """
    def __init__(self, dlist, flist = None):
        self.dlist = dlist
        self.flist = flist
        self.ind = 1
        self.val = ''
        self.fil = ''
        
    def vm(self):
        """ flags for new block && new vmerge val """
        f1=f2=0
        
        if self.dlist:
            if self.ind == self.dlist[0][0]:
                f1 = 1
                f2 = 1
            elif (self.ind > self.dlist[0][0]) & (self.ind < self.dlist[0][1]):
                f1 = 0
                f2 = 2
            elif self.ind == self.dlist[0][1]:
                f1 = 0
                f2 = 2
                self.dlist.pop(0)
            else:
                f1 = 1
                f2 = 0
        else:
            f1 = 1
            f2 = 0
            
        if f1:
            if self.flist:
                self.fil = self.flist.pop(0)
            else:
                self.fil = ''
                
        if f2 == 1:
            self.val = '<w:vMerge w:val="restart"'
        elif f2 == 2:
            self.val = '<w:vMerge w:val="continue"'
        else:
            self.val = ''
                        
        self.ind += 1
        
        return self.fil + self.val
