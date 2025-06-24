"""
Módulo para validação célula por célula com regras de tolerância
"""

import logging
from typing import Dict, Any, List, Tuple, Optional
import pandas as pd
import numpy as np
from decimal import Decimal, ROUND_HALF_UP


class Validator:
    """Validador de dados com suporte a diferentes regras de tolerância"""
    
    def __init__(self, tolerance_rules: Dict[str, Dict[str, Any]]):
        self.logger = logging.getLogger(__name__)
        self.tolerance_rules = tolerance_rules
        self.default_rule = tolerance_rules.get('default', {'tipo': 'exata'})
    
    def validate(self, df1: pd.DataFrame, df2: pd.DataFrame, 
                columns_to_validate: List[str], key_column: str = None) -> List[Dict[str, Any]]:
        """
        Valida dados célula por célula
        
        Args:
            df1: Primeiro DataFrame (fonte)
            df2: Segundo DataFrame (comparação)
            columns_to_validate: Lista de colunas para validar
            
        Returns:
            Lista de dicionários com resultados da validação
        """
        results = []
        
        # Validar que os DataFrames têm o mesmo número de linhas
        if len(df1) != len(df2):
            self.logger.warning(f"DataFrames têm tamanhos diferentes: {len(df1)} vs {len(df2)}")
        
        # Iterar sobre cada coluna a validar
        for col in columns_to_validate:
            self.logger.info(f"Validando coluna: {col}")
            
            # Verificar se a coluna existe em ambos os DataFrames
            if col not in df1.columns:
                self.logger.warning(f"Coluna '{col}' não encontrada no arquivo 1")
                continue
            if col not in df2.columns:
                self.logger.warning(f"Coluna '{col}' não encontrada no arquivo 2")
                continue
            
            # Obter regra de tolerância para esta coluna
            rule = self.tolerance_rules.get(col, self.default_rule)
            
            # Validar cada linha
            for idx in range(min(len(df1), len(df2))):
                val1 = df1.iloc[idx][col]
                val2 = df2.iloc[idx][col]
                
                # Comparar valores
                is_equal, tolerance_desc = self._compare_values(val1, val2, rule)
                
                # Adicionar resultado
                result = {
                    'linha_idx': idx,
                    'coluna': col,
                    'valor_arquivo_1': val1,
                    'valor_arquivo_2': val2,
                    'resultado': is_equal,
                    'tolerancia_aplicada': tolerance_desc
                }
                
                # Adicionar valor da coluna chave se disponível
                if key_column and key_column in df1.columns:
                    result['chave'] = df1.iloc[idx][key_column]
                
                results.append(result)
        
        # Log resumo
        total_validacoes = len(results)
        total_divergencias = sum(1 for r in results if not r['resultado'])
        self.logger.info(f"Validação concluída: {total_validacoes} comparações, "
                        f"{total_divergencias} divergências")
        
        return results
    
    def _compare_values(self, val1: Any, val2: Any, rule: Dict[str, Any]) -> Tuple[bool, str]:
        """
        Compara dois valores usando a regra especificada
        
        Args:
            val1: Primeiro valor
            val2: Segundo valor
            rule: Regra de tolerância
            
        Returns:
            Tupla (resultado_bool, descrição_tolerância)
        """
        # Tratar valores nulos
        if pd.isna(val1) and pd.isna(val2):
            return True, "Ambos nulos"
        if pd.isna(val1) or pd.isna(val2):
            return False, "Um valor nulo"
        
        rule_type = rule.get('tipo', 'exata')
        
        if rule_type == 'exata':
            # Comparação exata
            is_equal = val1 == val2
            return is_equal, "Exata"
        
        elif rule_type == 'decimal':
            # Comparação decimal com tolerância
            try:
                # Converter para float se necessário
                num1 = float(val1)
                num2 = float(val2)
                
                # Obter número de casas decimais
                decimal_places = rule.get('casas_decimais', 2)
                
                # Arredondar usando Decimal para precisão
                rounded1 = self._round_decimal(num1, decimal_places)
                rounded2 = self._round_decimal(num2, decimal_places)
                
                is_equal = rounded1 == rounded2
                return is_equal, f"Decimal, {decimal_places} casas"
                
            except (ValueError, TypeError):
                # Se não for possível converter para número, fazer comparação exata
                is_equal = val1 == val2
                return is_equal, "Exata (erro na conversão decimal)"
        
        else:
            # Tipo de regra desconhecido, usar comparação exata
            self.logger.warning(f"Tipo de regra desconhecido: {rule_type}")
            is_equal = val1 == val2
            return is_equal, "Exata (regra padrão)"
    
    def _round_decimal(self, value: float, decimal_places: int) -> Decimal:
        """
        Arredonda um valor para o número especificado de casas decimais
        usando arredondamento bancário (ROUND_HALF_UP)
        """
        decimal_value = Decimal(str(value))
        quantizer = Decimal('0.1') ** decimal_places
        return decimal_value.quantize(quantizer, rounding=ROUND_HALF_UP)
    
    def generate_summary(self, results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Gera um resumo dos resultados da validação
        
        Args:
            results: Lista de resultados da validação
            
        Returns:
            Dicionário com estatísticas resumidas
        """
        df_results = pd.DataFrame(results)
        
        summary = {
            'total_comparacoes': len(results),
            'total_iguais': len(df_results[df_results['resultado'] == True]),
            'total_divergencias': len(df_results[df_results['resultado'] == False]),
            'taxa_concordancia': 0.0,
            'divergencias_por_coluna': {},
            'tolerancias_aplicadas': {}
        }
        
        # Calcular taxa de concordância
        if summary['total_comparacoes'] > 0:
            summary['taxa_concordancia'] = (summary['total_iguais'] / summary['total_comparacoes']) * 100
        
        # Divergências por coluna
        for col in df_results['coluna'].unique():
            col_data = df_results[df_results['coluna'] == col]
            divergencias = len(col_data[col_data['resultado'] == False])
            summary['divergencias_por_coluna'][col] = {
                'total': len(col_data),
                'divergencias': divergencias,
                'taxa_divergencia': (divergencias / len(col_data)) * 100 if len(col_data) > 0 else 0
            }
        
        # Tolerâncias aplicadas
        for tol in df_results['tolerancia_aplicada'].unique():
            count = len(df_results[df_results['tolerancia_aplicada'] == tol])
            summary['tolerancias_aplicadas'][tol] = count
        
        return summary