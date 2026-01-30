#!/usr/bin/env python3
"""
Analyzes VBA forms to extract structure, dependencies, and functionality.
"""
import re
import os
from pathlib import Path
from typing import Dict, List, Set, Tuple

class VBAFormAnalyzer:
    def __init__(self, src_folder: str):
        self.src_folder = Path(src_folder)
        self.forms = {}
        self.modules = {}
        self.classes = {}
        
    def analyze_all_files(self) -> Dict:
        """Analyze all VBA files in the src folder."""
        results = {
            'forms': {},
            'modules': {},
            'classes': {},
            'dependencies': {},
            'unused_forms': set(),
            'dead_code_suspects': []
        }
        
        # Scan all files
        for file_path in self.src_folder.glob('**/*.cls'):
            if file_path.stem.startswith('Form_'):
                results['forms'][file_path.stem] = self.analyze_form(file_path)
            else:
                results['classes'][file_path.stem] = self.analyze_class(file_path)
                
        for file_path in self.src_folder.glob('**/*.bas'):
            results['modules'][file_path.stem] = self.analyze_module(file_path)
            
        # Analyze dependencies
        results['dependencies'] = self.extract_dependencies(results)
        
        # Detect unused forms
        results['unused_forms'] = self.detect_unused_forms(results)
        
        # Detect potential dead code
        results['dead_code_suspects'] = self.detect_dead_code(results)
        
        return results
    
    def analyze_form(self, file_path: Path) -> Dict:
        """Analyze a single form file."""
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        return {
            'name': file_path.stem,
            'path': str(file_path),
            'controls': self.extract_controls(content),
            'event_handlers': self.extract_event_handlers(content),
            'functions': self.extract_functions(content),
            'subroutines': self.extract_subroutines(content),
            'form_calls': self.extract_form_calls(content),
            'recordsources': self.extract_recordsources(content),
            'sql_queries': self.extract_sql_queries(content),
            'global_vars': self.extract_global_vars(content),
            'api_calls': self.extract_api_calls(content)
        }
    
    def analyze_module(self, file_path: Path) -> Dict:
        """Analyze a standard module (.bas)."""
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        return {
            'name': file_path.stem,
            'path': str(file_path),
            'functions': self.extract_functions(content),
            'subroutines': self.extract_subroutines(content),
            'global_vars': self.extract_global_vars(content),
            'constants': self.extract_constants(content),
            'is_utility': self.is_utility_module(content)
        }
    
    def analyze_class(self, file_path: Path) -> Dict:
        """Analyze a class module (.cls)."""
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        return {
            'name': file_path.stem,
            'path': str(file_path),
            'properties': self.extract_properties(content),
            'methods': self.extract_functions(content) + self.extract_subroutines(content),
            'events': self.extract_event_handlers(content)
        }
    
    def extract_controls(self, content: str) -> List[Dict]:
        """Extract form controls (textboxes, buttons, etc.)."""
        controls = []
        # Pattern for control declarations
        pattern = r'Begin\s+(\w+)\s+(\w+)'
        for match in re.finditer(pattern, content):
            control_type, control_name = match.groups()
            controls.append({
                'type': control_type,
                'name': control_name
            })
        return controls
    
    def extract_event_handlers(self, content: str) -> List[str]:
        """Extract event handler procedures."""
        pattern = r'(?:Private|Public)?\s*Sub\s+(\w+_\w+)\s*\('
        return [match.group(1) for match in re.finditer(pattern, content)]
    
    def extract_functions(self, content: str) -> List[Dict]:
        """Extract function definitions."""
        pattern = r'(?:Private|Public)?\s*Function\s+(\w+)\s*\((.*?)\)\s*As\s*(\w+)'
        functions = []
        for match in re.finditer(pattern, content):
            func_name, params, return_type = match.groups()
            functions.append({
                'name': func_name,
                'params': params.strip(),
                'return_type': return_type
            })
        return functions
    
    def extract_subroutines(self, content: str) -> List[Dict]:
        """Extract subroutine definitions."""
        pattern = r'(?:Private|Public)?\s*Sub\s+(\w+)\s*\((.*?)\)'
        subs = []
        for match in re.finditer(pattern, content):
            sub_name, params = match.groups()
            # Skip event handlers
            if '_' not in sub_name or not any(event in sub_name for event in ['Click', 'Load', 'Change', 'Enter', 'Exit']):
                subs.append({
                    'name': sub_name,
                    'params': params.strip()
                })
        return subs
    
    def extract_form_calls(self, content: str) -> List[str]:
        """Extract DoCmd.OpenForm calls to detect form navigation."""
        pattern = r'DoCmd\.OpenForm\s+["\']?([^"\'"\s,\)]+)'
        return list(set(match.group(1) for match in re.finditer(pattern, content)))
    
    def extract_recordsources(self, content: str) -> List[str]:
        """Extract RecordSource assignments."""
        pattern = r'(?:RecordSource|RowSource)\s*=\s*["\']([^"\']+)["\']'
        return list(set(match.group(1) for match in re.finditer(pattern, content)))
    
    def extract_sql_queries(self, content: str) -> List[str]:
        """Extract SQL query strings."""
        pattern = r'["\']SELECT\s+.+?FROM\s+.+?["\']'
        queries = []
        for match in re.finditer(pattern, content, re.IGNORECASE | re.DOTALL):
            query = match.group(0).strip('"\'')
            if len(query) < 500:  # Avoid extremely long queries
                queries.append(query)
        return queries[:10]  # Limit to first 10
    
    def extract_global_vars(self, content: str) -> List[Dict]:
        """Extract global/module-level variable declarations."""
        pattern = r'(?:Public|Private|Dim)\s+(\w+)\s+As\s+(\w+)'
        vars_dict = {}
        for match in re.finditer(pattern, content):
            var_name, var_type = match.groups()
            vars_dict[var_name] = var_type
        return [{'name': k, 'type': v} for k, v in vars_dict.items()]
    
    def extract_constants(self, content: str) -> List[Dict]:
        """Extract constant declarations."""
        pattern = r'(?:Public|Private)?\s*Const\s+(\w+)\s*(?:As\s+\w+)?\s*=\s*(.+?)(?:\r|\n)'
        constants = []
        for match in re.finditer(pattern, content):
            const_name, const_value = match.groups()
            constants.append({
                'name': const_name,
                'value': const_value.strip()
            })
        return constants
    
    def extract_properties(self, content: str) -> List[str]:
        """Extract property procedures from class modules."""
        pattern = r'(?:Public|Private)?\s*Property\s+(?:Get|Let|Set)\s+(\w+)'
        return list(set(match.group(1) for match in re.finditer(pattern, content)))
    
    def extract_api_calls(self, content: str) -> List[str]:
        """Extract external API/DLL calls."""
        pattern = r'Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+(\w+)\s+Lib'
        return [match.group(1) for match in re.finditer(pattern, content)]
    
    def is_utility_module(self, content: str) -> bool:
        """Determine if a module is a utility/helper module."""
        utility_keywords = ['Option Compare', 'Option Explicit', 'Public Function', 'Public Sub']
        return sum(1 for keyword in utility_keywords if keyword in content) >= 3
    
    def extract_dependencies(self, results: Dict) -> Dict:
        """Extract dependencies between forms and modules."""
        deps = {}
        
        for form_name, form_data in results['forms'].items():
            deps[form_name] = {
                'calls_forms': form_data['form_calls'],
                'uses_tables': self.extract_table_names(form_data['sql_queries'] + form_data['recordsources'])
            }
        
        return deps
    
    def extract_table_names(self, queries: List[str]) -> Set[str]:
        """Extract table names from SQL queries."""
        tables = set()
        for query in queries:
            # Simple pattern for FROM clause
            pattern = r'FROM\s+(\w+)'
            tables.update(match.group(1) for match in re.finditer(pattern, query, re.IGNORECASE))
        return list(tables)
    
    def detect_unused_forms(self, results: Dict) -> Set[str]:
        """Detect forms that are never opened by other forms."""
        all_forms = set(results['forms'].keys())
        called_forms = set()
        
        for form_data in results['forms'].values():
            called_forms.update(form_data['form_calls'])
        
        # Forms that are never called (except potential entry points)
        return all_forms - called_forms
    
    def detect_dead_code(self, results: Dict) -> List[Dict]:
        """Detect potential dead code (unused functions/subs)."""
        suspects = []
        
        # This is a simplified heuristic - would need more sophisticated analysis
        for form_name, form_data in results['forms'].items():
            for func in form_data['functions']:
                if func['name'].startswith('unused') or func['name'].startswith('old'):
                    suspects.append({
                        'location': form_name,
                        'type': 'function',
                        'name': func['name'],
                        'reason': 'Suspicious naming pattern'
                    })
        
        return suspects


if __name__ == '__main__':
    import sys
    import json
    
    if len(sys.argv) < 2:
        print("Usage: analyze_form.py <src_folder>")
        sys.exit(1)
    
    analyzer = VBAFormAnalyzer(sys.argv[1])
    results = analyzer.analyze_all_files()
    
    print(json.dumps(results, indent=2, default=str))
