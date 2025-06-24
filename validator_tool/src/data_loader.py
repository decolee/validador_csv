"""
Módulo para carregamento otimizado de arquivos CSV e XLSX
"""

import logging
from pathlib import Path
from typing import Optional, List, Union
import pandas as pd
import numpy as np


class DataLoader:
    """Carregador de dados com otimizações de memória"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def load_file(self, file_path: str, sheet_name: Optional[Union[str, int]] = None) -> pd.DataFrame:
        """
        Carrega arquivo CSV ou XLSX
        
        Args:
            file_path: Caminho para o arquivo
            sheet_name: Nome ou índice da aba (para XLSX)
            
        Returns:
            DataFrame com os dados
        """
        path = Path(file_path)
        
        if not path.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
        
        self.logger.info(f"Carregando arquivo: {file_path}")
        
        if path.suffix.lower() == '.csv':
            return self._load_csv(path)
        elif path.suffix.lower() in ['.xlsx', '.xls']:
            return self._load_excel(path, sheet_name)
        else:
            raise ValueError(f"Formato de arquivo não suportado: {path.suffix}")
    
    def load_excel_with_columns(self, file_path: str, sheet_name: Optional[Union[str, int]], 
                               columns_to_load: List[str]) -> pd.DataFrame:
        """
        Carrega apenas colunas específicas de um arquivo Excel (otimização de memória)
        
        Args:
            file_path: Caminho para o arquivo
            sheet_name: Nome ou índice da aba
            columns_to_load: Lista de colunas a carregar (nomes ou letras)
            
        Returns:
            DataFrame com apenas as colunas especificadas
        """
        path = Path(file_path)
        
        if not path.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
        
        self.logger.info(f"Carregando colunas específicas de: {file_path}")
        self.logger.debug(f"Colunas solicitadas: {columns_to_load}")
        
        # Converter letras de coluna para índices se necessário
        usecols = self._convert_column_letters_to_indices(columns_to_load)
        
        try:
            # Carregar apenas as colunas especificadas
            df = pd.read_excel(
                path,
                sheet_name=sheet_name,
                usecols=usecols,
                engine='openpyxl'
            )
            
            self.logger.info(f"Carregadas {len(df.columns)} colunas, {len(df)} linhas")
            return df
            
        except Exception as e:
            self.logger.error(f"Erro ao carregar arquivo Excel: {str(e)}")
            raise
    
    def _load_csv(self, path: Path) -> pd.DataFrame:
        """Carrega arquivo CSV com otimizações"""
        try:
            # Tentar detectar encoding
            encodings = ['utf-8', 'latin1', 'cp1252']
            
            for encoding in encodings:
                try:
                    # Tentar detectar o separador
                    df = pd.read_csv(
                        path,
                        encoding=encoding,
                        sep=';',  # Arquivos CSV usam ; como separador
                        decimal='.',  # Arquivos CSV usam . como decimal
                        low_memory=False,
                        na_values=['NA', 'N/A', 'null', 'NULL', '']
                    )
                    self.logger.info(f"CSV carregado com encoding: {encoding}")
                    return df
                except UnicodeDecodeError:
                    continue
            
            raise ValueError(f"Não foi possível detectar o encoding do arquivo: {path}")
            
        except Exception as e:
            self.logger.error(f"Erro ao carregar CSV: {str(e)}")
            raise
    
    def _load_excel(self, path: Path, sheet_name: Optional[Union[str, int]]) -> pd.DataFrame:
        """Carrega arquivo Excel"""
        try:
            df = pd.read_excel(
                path,
                sheet_name=sheet_name,
                engine='openpyxl',
                na_values=['NA', 'N/A', 'null', 'NULL', '']
            )
            
            if sheet_name is not None:
                self.logger.info(f"Excel carregado, aba: {sheet_name}")
            else:
                self.logger.info("Excel carregado, primeira aba")
                
            return df
            
        except Exception as e:
            self.logger.error(f"Erro ao carregar Excel: {str(e)}")
            raise
    
    def _convert_column_letters_to_indices(self, columns: List[str]) -> List[Union[str, int]]:
        """
        Converte letras de coluna (A, B, AA, etc) para índices ou mantém nomes
        """
        result = []
        
        for col in columns:
            # Se for uma letra de coluna Excel (A, B, AA, etc)
            if col.isalpha() and col.isupper():
                # Converter para índice (A=0, B=1, etc)
                index = 0
                for char in col:
                    index = index * 26 + (ord(char) - ord('A') + 1)
                result.append(index - 1)  # Ajustar para base 0
            else:
                # Manter como nome de coluna
                result.append(col)
        
        return result
    
    def optimize_dtypes(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Otimiza tipos de dados para reduzir uso de memória
        """
        self.logger.debug("Otimizando tipos de dados...")
        
        for col in df.columns:
            col_type = df[col].dtype
            
            # Otimizar inteiros
            if col_type != 'object':
                c_min = df[col].min()
                c_max = df[col].max()
                
                if str(col_type)[:3] == 'int':
                    if c_min > np.iinfo(np.int8).min and c_max < np.iinfo(np.int8).max:
                        df[col] = df[col].astype(np.int8)
                    elif c_min > np.iinfo(np.int16).min and c_max < np.iinfo(np.int16).max:
                        df[col] = df[col].astype(np.int16)
                    elif c_min > np.iinfo(np.int32).min and c_max < np.iinfo(np.int32).max:
                        df[col] = df[col].astype(np.int32)
                
                # Otimizar floats
                elif str(col_type)[:5] == 'float':
                    if c_min > np.finfo(np.float32).min and c_max < np.finfo(np.float32).max:
                        df[col] = df[col].astype(np.float32)
        
        return df