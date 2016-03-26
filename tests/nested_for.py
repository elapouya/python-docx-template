# -*- coding: utf-8 -*-
'''
Created : 2016-03-26

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

tpl=DocxTemplate('test_files/nested_for_tpl.docx')

context = {
    'dishes' : [
        {'name' : 'Pizza', 'ingredients' : ['bread','tomato', 'ham', 'cheese']},
        {'name' : 'Hamburger', 'ingredients' : ['bread','chopped steak', 'cheese', 'sauce']},
        {'name' : 'Apple pie', 'ingredients' : ['flour','apples', 'suggar', 'quince jelly']},
    ],
    'authors' : [
        {'name' : 'Saint-Exupery', 'books' : [
            {'title' : 'Le petit prince'},
            {'title' : "L'aviateur"},
            {'title' : 'Vol de nuit'},
        ]},
        {'name' : 'Barjavel', 'books' : [
            {'title' : 'Ravage'},
            {'title' : "La nuit des temps"},
            {'title' : 'Le grand secret'},
        ]},
    ]
}

tpl.render(context)
tpl.save('test_files/nested_for.docx')
