"""
Analisador de referências cruzadas entre abas (cross-sheet references)
Usa apenas bibliotecas nativas do Python para ambientes corporativos restritivos
"""

import re
import json
import csv
import logging
from typing import Dict, List, Set, Tuple, Any, Optional
from collections import defaultdict
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


class CrossSheetAnalyzer:
    """Analisa e mapeia referências entre diferentes abas de planilhas"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        # Padrões regex para diferentes tipos de referências
        self.patterns = {
            # Sheet!Cell - ex: Funcionarios!A1
            'sheet_cell': re.compile(r"([A-Za-z_][\w\s]*?)!([A-Z]+\d+)"),
            # Sheet!Range - ex: Dados!A1:B10
            'sheet_range': re.compile(r"([A-Za-z_][\w\s]*?)!([A-Z]+\d+:[A-Z]+\d+)"),
            # Sheet!Column - ex: Calculos!C:C
            'sheet_column': re.compile(r"([A-Za-z_][\w\s]*?)!([A-Z]+:[A-Z]+)"),
            # 'Sheet Name'!Cell - ex: 'Folha de Dados'!A1
            'quoted_sheet': re.compile(r"'([^']+)'!([A-Z]+\d+)"),
            # VLOOKUP com referência externa
            'vlookup_cross': re.compile(r"VLOOKUP\s*\([^,]+,\s*([A-Za-z_][\w\s]*?)!([^,\)]+)")
        }
    
    def analyze_workbook_native(self, file_path: str) -> Dict[str, Any]:
        """
        Analisa um arquivo Excel usando apenas bibliotecas nativas
        Extrai estrutura e fórmulas lendo o XML interno
        """
        workbook_data = {
            'sheets': {},
            'cross_references': defaultdict(list),
            'formulas': defaultdict(list),
            'dependencies': defaultdict(set)
        }
        
        try:
            with zipfile.ZipFile(file_path, 'r') as xlsx:
                # Ler relacionamentos de sheets
                sheet_info = self._extract_sheet_info(xlsx)
                
                # Processar cada sheet
                for sheet_name, sheet_file in sheet_info.items():
                    self.logger.info(f"Analisando aba: {sheet_name}")
                    
                    # Extrair dados da sheet
                    sheet_data = self._extract_sheet_data(xlsx, sheet_file)
                    workbook_data['sheets'][sheet_name] = sheet_data
                    
                    # Analisar fórmulas
                    for cell, formula in sheet_data.get('formulas', {}).items():
                        refs = self._extract_cross_references(formula, sheet_name)
                        if refs:
                            workbook_data['cross_references'][sheet_name].extend(refs)
                            # Mapear dependências
                            for ref in refs:
                                workbook_data['dependencies'][ref['target_sheet']].add(sheet_name)
        
        except Exception as e:
            self.logger.error(f"Erro ao analisar workbook: {str(e)}")
            raise
        
        return workbook_data
    
    def _extract_sheet_info(self, xlsx: zipfile.ZipFile) -> Dict[str, str]:
        """Extrai informações sobre as sheets do workbook"""
        sheet_info = {}
        
        # Ler workbook.xml para obter lista de sheets
        try:
            with xlsx.open('xl/workbook.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                # Namespace do Excel
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                
                sheets = root.findall('.//main:sheet', ns)
                for sheet in sheets:
                    name = sheet.get('name')
                    sheet_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    
                    # Mapear para arquivo físico
                    sheet_file = f"xl/worksheets/sheet{sheet_id[3:]}.xml"
                    sheet_info[name] = sheet_file
        
        except Exception as e:
            self.logger.error(f"Erro ao extrair informações de sheets: {str(e)}")
        
        return sheet_info
    
    def _extract_sheet_data(self, xlsx: zipfile.ZipFile, sheet_file: str) -> Dict[str, Any]:
        """Extrai dados e fórmulas de uma sheet específica"""
        sheet_data = {
            'cells': {},
            'formulas': {},
            'max_row': 0,
            'max_col': 0
        }
        
        try:
            with xlsx.open(sheet_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                
                # Processar células
                for cell in root.findall('.//main:c', ns):
                    cell_ref = cell.get('r')  # Ex: A1, B2
                    
                    # Extrair fórmula se existir
                    formula_elem = cell.find('main:f', ns)
                    if formula_elem is not None and formula_elem.text:
                        sheet_data['formulas'][cell_ref] = formula_elem.text
                    
                    # Extrair valor
                    value_elem = cell.find('main:v', ns)
                    if value_elem is not None and value_elem.text:
                        sheet_data['cells'][cell_ref] = value_elem.text
                    
                    # Atualizar dimensões
                    col, row = self._parse_cell_ref(cell_ref)
                    sheet_data['max_row'] = max(sheet_data['max_row'], row)
                    sheet_data['max_col'] = max(sheet_data['max_col'], col)
        
        except Exception as e:
            self.logger.error(f"Erro ao extrair dados da sheet: {str(e)}")
        
        return sheet_data
    
    def _extract_cross_references(self, formula: str, current_sheet: str) -> List[Dict[str, Any]]:
        """Extrai todas as referências cruzadas de uma fórmula"""
        references = []
        
        for pattern_name, pattern in self.patterns.items():
            matches = pattern.findall(formula)
            for match in matches:
                if isinstance(match, tuple):
                    target_sheet = match[0]
                    target_ref = match[1]
                else:
                    # Para padrões simples
                    parts = match.split('!')
                    if len(parts) == 2:
                        target_sheet = parts[0].strip("'")
                        target_ref = parts[1]
                    else:
                        continue
                
                if target_sheet != current_sheet:
                    references.append({
                        'type': pattern_name,
                        'source_sheet': current_sheet,
                        'target_sheet': target_sheet,
                        'target_reference': target_ref,
                        'formula': formula
                    })
        
        return references
    
    def _parse_cell_ref(self, cell_ref: str) -> Tuple[int, int]:
        """Converte referência de célula (A1) para coordenadas (col, row)"""
        match = re.match(r'([A-Z]+)(\d+)', cell_ref)
        if match:
            col_str = match.group(1)
            row = int(match.group(2))
            
            # Converter letras para número
            col = 0
            for char in col_str:
                col = col * 26 + (ord(char) - ord('A') + 1)
            
            return col, row
        return 0, 0
    
    def generate_dependency_graph(self, workbook_data: Dict[str, Any]) -> Dict[str, Any]:
        """Gera grafo de dependências entre sheets"""
        graph = {
            'nodes': [],
            'edges': [],
            'levels': {}
        }
        
        # Criar nós
        for sheet in workbook_data['sheets']:
            graph['nodes'].append({
                'id': sheet,
                'type': 'sheet',
                'formulas_count': len(workbook_data['sheets'][sheet].get('formulas', {}))
            })
        
        # Criar arestas baseadas em referências cruzadas
        for source_sheet, refs in workbook_data['cross_references'].items():
            targets = set()
            for ref in refs:
                targets.add(ref['target_sheet'])
            
            for target in targets:
                graph['edges'].append({
                    'source': source_sheet,
                    'target': target,
                    'count': len([r for r in refs if r['target_sheet'] == target])
                })
        
        # Calcular níveis (ordem de processamento)
        graph['levels'] = self._calculate_processing_order(
            workbook_data['dependencies']
        )
        
        return graph
    
    def _calculate_processing_order(self, dependencies: Dict[str, Set[str]]) -> Dict[int, List[str]]:
        """Calcula ordem de processamento baseada em dependências"""
        levels = {}
        processed = set()
        current_level = 0
        
        # Encontrar sheets sem dependências (nível 0)
        all_sheets = set()
        for sheet, deps in dependencies.items():
            all_sheets.add(sheet)
            all_sheets.update(deps)
        
        # Sheets que não dependem de ninguém
        level_0 = all_sheets - set(dependencies.keys())
        if level_0:
            levels[0] = list(level_0)
            processed.update(level_0)
            current_level = 1
        
        # Processar níveis subsequentes
        while len(processed) < len(all_sheets):
            current_level_sheets = []
            
            for sheet in all_sheets - processed:
                # Verificar se todas as dependências já foram processadas
                sheet_deps = dependencies.get(sheet, set())
                if sheet_deps.issubset(processed):
                    current_level_sheets.append(sheet)
            
            if current_level_sheets:
                levels[current_level] = current_level_sheets
                processed.update(current_level_sheets)
                current_level += 1
            else:
                # Detectar dependências circulares
                remaining = all_sheets - processed
                self.logger.warning(f"Possível dependência circular entre: {remaining}")
                levels[current_level] = list(remaining)
                break
        
        return levels
    
    def validate_cross_references(self, source_data: Dict[str, Any], 
                                target_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Valida se todas as referências cruzadas são válidas"""
        validation_errors = []
        
        for sheet, refs in source_data['cross_references'].items():
            for ref in refs:
                target_sheet = ref['target_sheet']
                target_ref = ref['target_reference']
                
                # Verificar se a sheet alvo existe
                if target_sheet not in target_data['sheets']:
                    validation_errors.append({
                        'type': 'missing_sheet',
                        'source_sheet': sheet,
                        'target_sheet': target_sheet,
                        'reference': target_ref,
                        'formula': ref['formula']
                    })
                    continue
                
                # Verificar se a célula/range existe
                if ':' not in target_ref:  # Célula única
                    if target_ref not in target_data['sheets'][target_sheet].get('cells', {}):
                        validation_errors.append({
                            'type': 'missing_cell',
                            'source_sheet': sheet,
                            'target_sheet': target_sheet,
                            'reference': target_ref,
                            'formula': ref['formula']
                        })
        
        return validation_errors
    
    def suggest_optimizations(self, workbook_data: Dict[str, Any]) -> List[Dict[str, str]]:
        """Sugere otimizações baseadas na análise de dependências"""
        suggestions = []
        
        # Detectar dependências circulares
        circular = self._detect_circular_dependencies(workbook_data['dependencies'])
        if circular:
            suggestions.append({
                'tipo': 'dependencia_circular',
                'descricao': f"Dependências circulares detectadas: {' -> '.join(circular)}",
                'impacto': 'alto',
                'solucao': 'Reestruturar cálculos para eliminar referências circulares'
            })
        
        # Detectar sheets com muitas dependências
        for sheet, deps in workbook_data['dependencies'].items():
            if len(deps) > 5:
                suggestions.append({
                    'tipo': 'alta_complexidade',
                    'descricao': f"Sheet '{sheet}' tem {len(deps)} dependências",
                    'impacto': 'medio',
                    'solucao': 'Considerar consolidar ou simplificar cálculos'
                })
        
        # Detectar fórmulas com muitas referências externas
        for sheet, refs in workbook_data['cross_references'].items():
            if len(refs) > 20:
                suggestions.append({
                    'tipo': 'excesso_referencias',
                    'descricao': f"Sheet '{sheet}' tem {len(refs)} referências externas",
                    'impacto': 'medio',
                    'solucao': 'Considerar usar tabelas auxiliares locais'
                })
        
        return suggestions
    
    def _detect_circular_dependencies(self, dependencies: Dict[str, Set[str]]) -> List[str]:
        """Detecta dependências circulares usando DFS"""
        def has_cycle(node, visited, rec_stack, path):
            visited[node] = True
            rec_stack[node] = True
            path.append(node)
            
            for neighbor in dependencies.get(node, []):
                if not visited.get(neighbor, False):
                    if has_cycle(neighbor, visited, rec_stack, path):
                        return True
                elif rec_stack.get(neighbor, False):
                    # Encontrou ciclo
                    cycle_start = path.index(neighbor)
                    return path[cycle_start:]
            
            path.pop()
            rec_stack[node] = False
            return False
        
        visited = {}
        rec_stack = {}
        
        for node in dependencies:
            if not visited.get(node, False):
                path = []
                result = has_cycle(node, visited, rec_stack, path)
                if result and isinstance(result, list):
                    return result
        
        return []