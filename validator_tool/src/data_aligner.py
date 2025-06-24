"""
Módulo para alinhamento de dados usando coluna chave
"""

import logging
from typing import Dict, List, Tuple
import pandas as pd


class DataAligner:
    """Alinha DataFrames usando colunas chave"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def align_dataframes(self, df1: pd.DataFrame, df2: pd.DataFrame, 
                        key_col1: str, key_col2: str) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, List]]:
        """
        Alinha dois DataFrames usando colunas chave
        
        Args:
            df1: Primeiro DataFrame
            df2: Segundo DataFrame
            key_col1: Nome da coluna chave no df1
            key_col2: Nome da coluna chave no df2
            
        Returns:
            Tupla com (df1_alinhado, df2_alinhado, dict_linhas_sem_correspondencia)
        """
        # Validar colunas chave
        if key_col1 not in df1.columns:
            raise ValueError(f"Coluna chave '{key_col1}' não encontrada no primeiro arquivo")
        if key_col2 not in df2.columns:
            raise ValueError(f"Coluna chave '{key_col2}' não encontrada no segundo arquivo")
        
        # Criar cópias para não modificar originais
        df1_work = df1.copy()
        df2_work = df2.copy()
        
        # Adicionar índice original para rastreamento
        df1_work['_original_index_1'] = df1_work.index
        df2_work['_original_index_2'] = df2_work.index
        
        # Renomear colunas chave para facilitar merge
        df1_work = df1_work.rename(columns={key_col1: '_merge_key'})
        df2_work = df2_work.rename(columns={key_col2: '_merge_key'})
        
        # Realizar merge
        self.logger.info("Realizando merge dos DataFrames...")
        merged = pd.merge(
            df1_work,
            df2_work,
            on='_merge_key',
            how='outer',
            indicator=True,
            suffixes=('_df1', '_df2')
        )
        
        # Separar registros alinhados e não alinhados
        aligned = merged[merged['_merge'] == 'both'].copy()
        only_df1 = merged[merged['_merge'] == 'left_only'].copy()
        only_df2 = merged[merged['_merge'] == 'right_only'].copy()
        
        # Preparar DataFrames alinhados
        if len(aligned) > 0:
            # Reconstruir df1 alinhado
            df1_cols = [col for col in aligned.columns if col.endswith('_df1') or col == '_merge_key']
            df1_aligned = aligned[df1_cols].copy()
            
            # Remover sufixos e colunas auxiliares
            df1_aligned.columns = [col.replace('_df1', '') if col != '_merge_key' else key_col1 
                                 for col in df1_aligned.columns]
            df1_aligned = df1_aligned.drop(columns=['_original_index_1'], errors='ignore')
            
            # Reconstruir df2 alinhado
            df2_cols = [col for col in aligned.columns if col.endswith('_df2') or col == '_merge_key']
            df2_aligned = aligned[df2_cols].copy()
            
            # Remover sufixos e colunas auxiliares
            df2_aligned.columns = [col.replace('_df2', '') if col != '_merge_key' else key_col2 
                                 for col in df2_aligned.columns]
            df2_aligned = df2_aligned.drop(columns=['_original_index_2'], errors='ignore')
            
            # Resetar índices
            df1_aligned = df1_aligned.reset_index(drop=True)
            df2_aligned = df2_aligned.reset_index(drop=True)
        else:
            # Nenhuma correspondência encontrada
            df1_aligned = pd.DataFrame()
            df2_aligned = pd.DataFrame()
        
        # Coletar linhas sem correspondência
        unmatched = {
            'df1': [],
            'df2': []
        }
        
        if len(only_df1) > 0:
            unmatched['df1'] = only_df1['_merge_key'].tolist()
        
        if len(only_df2) > 0:
            unmatched['df2'] = only_df2['_merge_key'].tolist()
        
        # Log estatísticas
        self.logger.info(f"Alinhamento concluído:")
        self.logger.info(f"  - Linhas alinhadas: {len(aligned)}")
        self.logger.info(f"  - Linhas apenas no arquivo 1: {len(unmatched['df1'])}")
        self.logger.info(f"  - Linhas apenas no arquivo 2: {len(unmatched['df2'])}")
        
        return df1_aligned, df2_aligned, unmatched
    
    def create_unmatched_report(self, unmatched: Dict[str, List]) -> pd.DataFrame:
        """
        Cria um DataFrame com relatório de linhas sem correspondência
        
        Args:
            unmatched: Dicionário com listas de valores chave sem correspondência
            
        Returns:
            DataFrame com o relatório
        """
        records = []
        
        # Linhas apenas no arquivo 1
        for key in unmatched['df1']:
            records.append({
                'Chave': key,
                'Presente_em': 'Apenas Arquivo 1',
                'Observacao': 'Sem correspondência no Arquivo 2'
            })
        
        # Linhas apenas no arquivo 2
        for key in unmatched['df2']:
            records.append({
                'Chave': key,
                'Presente_em': 'Apenas Arquivo 2',
                'Observacao': 'Sem correspondência no Arquivo 1'
            })
        
        return pd.DataFrame(records)