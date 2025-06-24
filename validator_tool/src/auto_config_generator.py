"""
Gerador automático de configuração - Analisa arquivos e sugere config.json
Usa apenas bibliotecas nativas do Python
"""

import json
import csv
import re
import logging
from typing import Dict, List, Any, Optional, Tuple, Set
from pathlib import Path
from collections import defaultdict, Counter
import statistics
import zipfile
import xml.etree.ElementTree as ET


class AutoConfigGenerator:
    """Gera configuração automaticamente analisando os arquivos"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.numeric_pattern = re.compile(r'^-?\d+\.?\d*$')
        self.date_patterns = [
            re.compile(r'^\d{4}-\d{2}-\d{2}$'),  # YYYY-MM-DD
            re.compile(r'^\d{2}/\d{2}/\d{4}$'),  # DD/MM/YYYY
            re.compile(r'^\d{2}-\d{2}-\d{4}$'),  # DD-MM-YYYY
        ]
    
    def analyze_files(self, file_paths: List[str]) -> Dict[str, Any]:
        """
        Analisa múltiplos arquivos e gera configuração sugerida
        """
        analysis = {
            'files': {},
            'common_columns': set(),
            'suggested_keys': [],
            'suggested_validations': [],
            'suggested_tolerances': {},
            'relationships': [],
            'formulas_detected': {}
        }
        
        # Analisar cada arquivo
        for file_path in file_paths:
            self.logger.info(f"Analisando arquivo: {file_path}")
            file_analysis = self._analyze_single_file(file_path)
            analysis['files'][file_path] = file_analysis
        
        # Encontrar elementos comuns e sugestões
        self._find_common_elements(analysis)
        self._suggest_key_columns(analysis)
        self._suggest_validation_columns(analysis)
        self._suggest_tolerance_rules(analysis)
        self._detect_relationships(analysis)
        
        return analysis
    
    def _analyze_single_file(self, file_path: str) -> Dict[str, Any]:
        """Analisa um único arquivo"""
        path = Path(file_path)
        
        if path.suffix.lower() == '.csv':
            return self._analyze_csv(file_path)
        elif path.suffix.lower() in ['.xlsx', '.xls']:
            return self._analyze_excel(file_path)
        else:
            raise ValueError(f"Formato não suportado: {path.suffix}")
    
    def _analyze_csv(self, file_path: str) -> Dict[str, Any]:
        """Analisa arquivo CSV"""
        analysis = {
            'format': 'csv',
            'sheets': {},
            'encoding': None,
            'delimiter': None,
            'columns': {},
            'row_count': 0
        }
        
        # Detectar encoding e delimiter
        encodings = ['utf-8', 'latin1', 'cp1252']
        delimiters = [',', ';', '\t', '|']
        
        for encoding in encodings:
            for delimiter in delimiters:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        # Tentar ler algumas linhas
                        sample = []
                        reader = csv.reader(f, delimiter=delimiter)
                        for i, row in enumerate(reader):
                            sample.append(row)
                            if i >= 10:  # Amostra de 10 linhas
                                break
                        
                        # Verificar se parece válido
                        if sample and len(sample[0]) > 1 and all(len(row) == len(sample[0]) for row in sample[1:]):
                            analysis['encoding'] = encoding
                            analysis['delimiter'] = delimiter
                            
                            # Analisar estrutura completa
                            f.seek(0)
                            self._analyze_csv_structure(f, reader, analysis)
                            return analysis
                
                except:
                    continue
        
        raise ValueError(f"Não foi possível detectar formato do CSV: {file_path}")
    
    def _analyze_csv_structure(self, file_handle, reader, analysis: Dict[str, Any]):
        """Analisa estrutura do CSV"""
        headers = next(reader)
        analysis['sheets']['main'] = {
            'columns': headers,
            'column_analysis': {}
        }
        
        # Analisar tipos de dados
        data_samples = defaultdict(list)
        row_count = 0
        
        for row in reader:
            row_count += 1
            for i, value in enumerate(row):
                if i < len(headers):
                    data_samples[headers[i]].append(value)
        
        analysis['row_count'] = row_count
        
        # Analisar cada coluna
        for col, values in data_samples.items():
            analysis['columns'][col] = self._analyze_column_data(values)
            analysis['sheets']['main']['column_analysis'][col] = analysis['columns'][col]
    
    def _analyze_excel(self, file_path: str) -> Dict[str, Any]:
        """Analisa arquivo Excel usando apenas bibliotecas nativas"""
        analysis = {
            'format': 'excel',
            'sheets': {},
            'columns': {},
            'formulas': {},
            'relationships': []
        }
        
        try:
            with zipfile.ZipFile(file_path, 'r') as xlsx:
                # Listar sheets
                sheet_info = self._get_excel_sheets(xlsx)
                
                for sheet_name, sheet_file in sheet_info.items():
                    self.logger.debug(f"Analisando sheet: {sheet_name}")
                    sheet_analysis = self._analyze_excel_sheet(xlsx, sheet_file)
                    analysis['sheets'][sheet_name] = sheet_analysis
                    
                    # Consolidar análise de colunas
                    for col, col_analysis in sheet_analysis.get('column_analysis', {}).items():
                        full_col_name = f"{sheet_name}.{col}"
                        analysis['columns'][full_col_name] = col_analysis
        
        except Exception as e:
            self.logger.error(f"Erro ao analisar Excel: {str(e)}")
            raise
        
        return analysis
    
    def _get_excel_sheets(self, xlsx: zipfile.ZipFile) -> Dict[str, str]:
        """Obtém lista de sheets do Excel"""
        sheets = {}
        
        with xlsx.open('xl/workbook.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            for sheet in root.findall('.//main:sheet', ns):
                name = sheet.get('name')
                sheet_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                sheet_file = f"xl/worksheets/sheet{sheet_id[3:]}.xml"
                sheets[name] = sheet_file
        
        return sheets
    
    def _analyze_excel_sheet(self, xlsx: zipfile.ZipFile, sheet_file: str) -> Dict[str, Any]:
        """Analisa uma sheet específica do Excel"""
        analysis = {
            'columns': [],
            'column_analysis': {},
            'formulas_count': 0,
            'row_count': 0
        }
        
        with xlsx.open(sheet_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()
            
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            # Coletar dados por coluna
            column_data = defaultdict(list)
            formulas = defaultdict(list)
            
            for cell in root.findall('.//main:c', ns):
                cell_ref = cell.get('r')
                col_letter = re.match(r'([A-Z]+)', cell_ref).group(1)
                
                # Verificar se tem fórmula
                formula = cell.find('main:f', ns)
                if formula is not None and formula.text:
                    formulas[col_letter].append(formula.text)
                    analysis['formulas_count'] += 1
                
                # Obter valor
                value = cell.find('main:v', ns)
                if value is not None and value.text:
                    column_data[col_letter].append(value.text)
            
            # Analisar cada coluna
            for col, values in column_data.items():
                if values:  # Apenas colunas com dados
                    analysis['columns'].append(col)
                    analysis['column_analysis'][col] = self._analyze_column_data(values)
                    if col in formulas:
                        analysis['column_analysis'][col]['has_formulas'] = True
                        analysis['column_analysis'][col]['formula_sample'] = formulas[col][0]
            
            analysis['row_count'] = len(next(iter(column_data.values()), []))
        
        return analysis
    
    def _analyze_column_data(self, values: List[str]) -> Dict[str, Any]:
        """Analisa dados de uma coluna para determinar tipo e características"""
        analysis = {
            'type': 'unknown',
            'nullable': False,
            'unique_count': 0,
            'null_count': 0,
            'patterns': [],
            'statistics': {}
        }
        
        # Filtrar valores vazios
        non_empty = [v for v in values if v and str(v).strip()]
        analysis['null_count'] = len(values) - len(non_empty)
        analysis['nullable'] = analysis['null_count'] > 0
        
        if not non_empty:
            analysis['type'] = 'empty'
            return analysis
        
        # Contar valores únicos
        analysis['unique_count'] = len(set(non_empty))
        analysis['uniqueness_ratio'] = analysis['unique_count'] / len(non_empty)
        
        # Detectar tipo
        numeric_count = sum(1 for v in non_empty if self._is_numeric(v))
        date_count = sum(1 for v in non_empty if self._is_date(v))
        
        if numeric_count == len(non_empty):
            analysis['type'] = 'numeric'
            # Estatísticas numéricas
            numbers = [float(v) for v in non_empty]
            analysis['statistics'] = {
                'min': min(numbers),
                'max': max(numbers),
                'mean': statistics.mean(numbers),
                'stdev': statistics.stdev(numbers) if len(numbers) > 1 else 0,
                'decimals': max(len(str(n).split('.')[-1]) for n in numbers if '.' in str(n))
            }
        elif date_count == len(non_empty):
            analysis['type'] = 'date'
        elif analysis['uniqueness_ratio'] > 0.9:
            analysis['type'] = 'identifier'
        else:
            analysis['type'] = 'text'
            # Detectar padrões comuns
            if len(set(non_empty)) < 10:
                analysis['patterns'] = list(set(non_empty))
        
        return analysis
    
    def _is_numeric(self, value: str) -> bool:
        """Verifica se o valor é numérico"""
        try:
            float(value.replace(',', '.'))
            return True
        except:
            return False
    
    def _is_date(self, value: str) -> bool:
        """Verifica se o valor é data"""
        return any(pattern.match(str(value)) for pattern in self.date_patterns)
    
    def _find_common_elements(self, analysis: Dict[str, Any]):
        """Encontra elementos comuns entre arquivos"""
        all_columns = []
        
        for file_data in analysis['files'].values():
            if file_data['format'] == 'csv':
                all_columns.extend(file_data['sheets']['main']['columns'])
            else:
                for sheet_data in file_data['sheets'].values():
                    all_columns.extend(sheet_data['columns'])
        
        # Encontrar colunas que aparecem em múltiplos arquivos
        column_counts = Counter(all_columns)
        analysis['common_columns'] = {col for col, count in column_counts.items() if count > 1}
    
    def _suggest_key_columns(self, analysis: Dict[str, Any]):
        """Sugere colunas chave baseado em características"""
        candidates = []
        
        for file_path, file_data in analysis['files'].items():
            for col_name, col_analysis in file_data['columns'].items():
                if col_analysis['type'] == 'identifier' or (
                    col_analysis['uniqueness_ratio'] > 0.95 and 
                    col_analysis['null_count'] == 0
                ):
                    candidates.append({
                        'column': col_name,
                        'file': file_path,
                        'score': col_analysis['uniqueness_ratio']
                    })
        
        # Ordenar por score
        candidates.sort(key=lambda x: x['score'], reverse=True)
        analysis['suggested_keys'] = candidates[:5]  # Top 5 candidatos
    
    def _suggest_validation_columns(self, analysis: Dict[str, Any]):
        """Sugere colunas para validação baseado em importância"""
        suggestions = []
        
        # Priorizar colunas numéricas com fórmulas
        for file_path, file_data in analysis['files'].items():
            for col_name, col_analysis in file_data['columns'].items():
                score = 0
                reasons = []
                
                if col_analysis['type'] == 'numeric':
                    score += 2
                    reasons.append('numeric')
                
                if col_analysis.get('has_formulas'):
                    score += 3
                    reasons.append('has_formulas')
                
                if col_name in analysis['common_columns']:
                    score += 1
                    reasons.append('common_column')
                
                if score > 0:
                    suggestions.append({
                        'column': col_name,
                        'score': score,
                        'reasons': reasons
                    })
        
        # Ordenar por score
        suggestions.sort(key=lambda x: x['score'], reverse=True)
        analysis['suggested_validations'] = [s['column'] for s in suggestions[:20]]
    
    def _suggest_tolerance_rules(self, analysis: Dict[str, Any]):
        """Sugere regras de tolerância baseado nos tipos de dados"""
        rules = {}
        
        for file_path, file_data in analysis['files'].items():
            for col_name, col_analysis in file_data['columns'].items():
                if col_analysis['type'] == 'numeric':
                    decimals = col_analysis.get('statistics', {}).get('decimals', 0)
                    if decimals > 0:
                        rules[col_name] = {
                            'tipo': 'decimal',
                            'casas_decimais': min(decimals, 10)
                        }
                    else:
                        rules[col_name] = {'tipo': 'exata'}
                elif col_analysis['type'] == 'date':
                    rules[col_name] = {'tipo': 'data'}
                else:
                    rules[col_name] = {'tipo': 'exata'}
        
        rules['default'] = {'tipo': 'exata'}
        analysis['suggested_tolerances'] = rules
    
    def _detect_relationships(self, analysis: Dict[str, Any]):
        """Detecta possíveis relacionamentos entre arquivos/abas"""
        relationships = []
        
        # Procurar por colunas com nomes similares
        column_groups = defaultdict(list)
        
        for file_path, file_data in analysis['files'].items():
            for col_name in file_data['columns']:
                # Normalizar nome da coluna
                normalized = re.sub(r'[_\s-]', '', col_name.upper())
                column_groups[normalized].append({
                    'file': file_path,
                    'column': col_name
                })
        
        # Identificar possíveis relacionamentos
        for normalized, occurrences in column_groups.items():
            if len(occurrences) > 1:
                relationships.append({
                    'type': 'possible_join',
                    'columns': occurrences,
                    'confidence': 'high' if 'ID' in normalized else 'medium'
                })
        
        analysis['relationships'] = relationships
    
    def generate_config(self, analysis: Dict[str, Any], output_path: Optional[str] = None) -> Dict[str, Any]:
        """Gera arquivo de configuração baseado na análise"""
        config = {
            'gerado_automaticamente': True,
            'versao': '2.0',
            'arquivos': {},
            'validacoes': {},
            'configuracoes_avancadas': {}
        }
        
        # Mapear arquivos
        file_list = list(analysis['files'].keys())
        if len(file_list) >= 2:
            config['arquivo_fonte_1'] = {
                'caminho': file_list[0],
                'formato_detectado': analysis['files'][file_list[0]]['format']
            }
            config['arquivo_fonte_2'] = {
                'caminho': file_list[1],
                'formato_detectado': analysis['files'][file_list[1]]['format']
            }
            
            # Adicionar detalhes específicos do formato
            for i, file_path in enumerate(file_list[:2], 1):
                file_key = f'arquivo_fonte_{i}'
                file_data = analysis['files'][file_path]
                
                if file_data['format'] == 'csv':
                    config[file_key]['encoding'] = file_data.get('encoding', 'utf-8')
                    config[file_key]['delimiter'] = file_data.get('delimiter', ',')
                else:
                    # Para Excel, listar abas disponíveis
                    config[file_key]['abas_disponiveis'] = list(file_data['sheets'].keys())
        
        # Sugerir coluna chave
        if analysis['suggested_keys']:
            key_suggestion = analysis['suggested_keys'][0]['column']
            # Remover prefixo de sheet se houver
            if '.' in key_suggestion:
                key_suggestion = key_suggestion.split('.')[-1]
            
            config['arquivo_fonte_1']['coluna_chave'] = key_suggestion
            config['arquivo_fonte_2']['coluna_chave'] = key_suggestion
        
        # Colunas para validar
        config['colunas_para_validar'] = analysis['suggested_validations'][:10]
        
        # Regras de tolerância
        config['regras_de_tolerancia'] = analysis['suggested_tolerances']
        
        # Adicionar seção de relacionamentos detectados
        if analysis['relationships']:
            config['relacionamentos_detectados'] = analysis['relationships']
        
        # Salvar arquivo se caminho fornecido
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            self.logger.info(f"Configuração salva em: {output_path}")
        
        return config