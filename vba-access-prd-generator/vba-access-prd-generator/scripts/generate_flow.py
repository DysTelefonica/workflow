#!/usr/bin/env python3
"""
Generates a navigation flow diagram from form dependencies.
"""
import json
import sys
from pathlib import Path

def generate_mermaid_flow(dependencies: dict, forms: dict) -> str:
    """Generate a Mermaid flowchart from form dependencies."""
    
    mermaid = ["graph TD"]
    
    # Add nodes
    for form_name in forms.keys():
        clean_name = form_name.replace('Form_', '')
        mermaid.append(f'    {clean_name}["{clean_name}"]')
    
    # Add edges
    for form_name, deps in dependencies.items():
        clean_source = form_name.replace('Form_', '')
        for called_form in deps.get('calls_forms', []):
            clean_target = called_form.replace('Form_', '')
            mermaid.append(f'    {clean_source} --> {clean_target}')
    
    return '\n'.join(mermaid)

def generate_module_dependencies(forms: dict, modules: dict) -> str:
    """Generate a dependency diagram including modules."""
    
    mermaid = ["graph LR"]
    mermaid.append("    subgraph Forms")
    
    for form_name in forms.keys():
        clean_name = form_name.replace('Form_', '')
        mermaid.append(f'        {clean_name}["{clean_name}"]')
    
    mermaid.append("    end")
    mermaid.append("    subgraph Modules")
    
    for module_name in modules.keys():
        mermaid.append(f'        {module_name}["{module_name}"]')
    
    mermaid.append("    end")
    
    return '\n'.join(mermaid)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: generate_flow.py <analysis_json_file>")
        sys.exit(1)
    
    with open(sys.argv[1], 'r') as f:
        data = json.load(f)
    
    print("# Navigation Flow Diagram\n")
    print("```mermaid")
    print(generate_mermaid_flow(data['dependencies'], data['forms']))
    print("```\n")
    
    print("# Module Dependencies\n")
    print("```mermaid")
    print(generate_module_dependencies(data['forms'], data['modules']))
    print("```")
