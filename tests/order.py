# -*- coding: utf-8 -*-
'''
Created : 2015-03-12

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

tpl=DocxTemplate('test_files/order_tpl.docx')

context = { 
    'customer_name' : 'Eric',
    'items' : [
        {'desc' : 'Python interpreters', 'qty' : 2, 'price' : 'FREE' },
        {'desc' : 'Django projects', 'qty' : 5403, 'price' : 'FREE' },
        {'desc' : 'Guido', 'qty' : 1, 'price' : '100,000,000.00' },
    ], 
    'in_europe' : True, 
    'is_paid': False,
    'company_name' : 'The World Wide company', 
    'total_price' : '100,000,000.00' 
}

tpl.render(context)
tpl.save('test_files/order.docx')
