#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test for get_undeclared_template_variables method

This test demonstrates the correct behavior of get_undeclared_template_variables:
1. Before rendering - finds all template variables
2. After rendering with incomplete context - finds missing variables
3. After rendering with complete context - returns empty set
"""

from docxtpl import DocxTemplate

def test_before_render():
    """Test that get_undeclared_template_variables finds all variables before rendering"""
    print("=== Test 1: Before render ===")
    tpl = DocxTemplate('templates/get_undeclared_variables.docx')
    undeclared = tpl.get_undeclared_template_variables()
    print(f"Variables found: {undeclared}")
    
    # Should find all variables
    expected_vars = {
        'name', 'age', 'email', 'is_student', 'has_degree', 'degree_field',
        'skills', 'projects', 'company_name', 'page_number', 'generation_date', 'author'
    }
    
    if undeclared == expected_vars:
        print("PASS: Found all expected variables before render")
    else:
        print(f"FAIL: Expected {expected_vars}, got {undeclared}")
    
    return undeclared == expected_vars

def test_after_incomplete_render():
    """Test that get_undeclared_template_variables finds missing variables after incomplete render"""
    print("\n=== Test 2: After incomplete render ===")
    tpl = DocxTemplate('templates/get_undeclared_variables.docx')
    
    # Provide only some variables (missing several)
    context = {
        'name': 'John Doe',
        'age': 25,
        'email': 'john@example.com',
        'is_student': True,
        'skills': ['Python', 'Django'],
        'company_name': 'Test Corp',
        'author': 'Test Author'
    }
    
    tpl.render(context)
    undeclared = tpl.get_undeclared_template_variables(context=context)
    print(f"Missing variables: {undeclared}")
    
    # Should find missing variables
    expected_missing = {
        'has_degree', 'degree_field', 'projects', 'page_number', 'generation_date'
    }
    
    if undeclared == expected_missing:
        print("PASS: Found missing variables after incomplete render")
    else:
        print(f"FAIL: Expected missing {expected_missing}, got {undeclared}")
    
    return undeclared == expected_missing

def test_after_complete_render():
    """Test that get_undeclared_template_variables returns empty set after complete render"""
    print("\n=== Test 3: After complete render ===")
    tpl = DocxTemplate('templates/get_undeclared_variables.docx')
    
    # Provide all variables
    context = {
        'name': 'John Doe',
        'age': 25,
        'email': 'john@example.com',
        'is_student': True,
        'has_degree': True,
        'degree_field': 'Computer Science',
        'skills': ['Python', 'Django', 'JavaScript'],
        'projects': [
            {'name': 'Project A', 'year': 2023, 'description': 'A great project'},
            {'name': 'Project B', 'year': 2024, 'description': 'Another great project'}
        ],
        'company_name': 'Test Corp',
        'page_number': 1,
        'generation_date': '2024-01-15',
        'author': 'Test Author'
    }
    
    tpl.render(context)
    undeclared = tpl.get_undeclared_template_variables(context=context)
    print(f"Undeclared variables: {undeclared}")
    
    # Should return empty set
    if undeclared == set():
        print("PASS: No undeclared variables after complete render")
    else:
        print(f"FAIL: Expected empty set, got {undeclared}")
    
    return undeclared == set()

def test_with_custom_jinja_env():
    """Test that get_undeclared_template_variables works with custom Jinja environment"""
    print("\n=== Test 4: With custom Jinja environment ===")
    from jinja2 import Environment
    
    tpl = DocxTemplate('templates/get_undeclared_variables.docx')
    custom_env = Environment()
    
    undeclared = tpl.get_undeclared_template_variables(jinja_env=custom_env)
    print(f"Variables found with custom env: {undeclared}")
    
    # Should find all variables
    expected_vars = {
        'name', 'age', 'email', 'is_student', 'has_degree', 'degree_field',
        'skills', 'projects', 'company_name', 'page_number', 'generation_date', 'author'
    }
    
    if undeclared == expected_vars:
        print("PASS: Custom Jinja environment works correctly")
    else:
        print(f"FAIL: Expected {expected_vars}, got {undeclared}")
    
    return undeclared == expected_vars

if __name__ == "__main__":
    print("Testing get_undeclared_template_variables method...")
    print("=" * 50)
    
    # Run all tests
    test1_passed = test_before_render()
    test2_passed = test_after_incomplete_render()
    test3_passed = test_after_complete_render()
    test4_passed = test_with_custom_jinja_env()
    
    print("\n" + "=" * 50)
    print("SUMMARY:")
    print(f"Test 1 (Before render): {'PASS' if test1_passed else 'FAIL'}")
    print(f"Test 2 (After incomplete render): {'PASS' if test2_passed else 'FAIL'}")
    print(f"Test 3 (After complete render): {'PASS' if test3_passed else 'FAIL'}")
    print(f"Test 4 (Custom Jinja env): {'PASS' if test4_passed else 'FAIL'}")
    
    all_passed = test1_passed and test2_passed and test3_passed and test4_passed
    
    if all_passed:
        print("ALL TESTS PASSED!")
    else:
        print("SOME TESTS FAILED!")
    
    print("\nNote: This test demonstrates that get_undeclared_template_variables")
    print("now correctly analyzes the original template, not the rendered document.")