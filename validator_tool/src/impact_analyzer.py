"""
Módulo para análise de impacto de divergências em fórmulas - Versão 2
"""

import logging
import re
from typing import Dict, Any, List, Set

class ImpactAnalyzer:
    """Analisa o impacto de divergências nas fórmulas"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def analyze_formula_dependencies(self, validation_results: List[Dict[str, Any]], 
                                   formulas: Dict[str, Dict[str, Any]]) -> Dict[str, List[str]]:
        """
        Analisa quais colunas divergentes impactam cada fórmula
        
        Para cada resultado de validação de uma coluna com fórmula,
        identifica quais colunas base (referenciadas na fórmula) estão
        causando a divergência.
        """
        # 1. Construir mapa de dependências inversas
        # col_base -> [formulas que dependem dela]
        dependencias_inversas = {}
        
        for col_formula, formula_info in formulas.items():
            mapeamento = formula_info.get('mapeamento', {})
            if isinstance(mapeamento, dict):
                for col_referenciada in mapeamento.values():
                    if col_referenciada not in dependencias_inversas:
                        dependencias_inversas[col_referenciada] = []
                    dependencias_inversas[col_referenciada].append(col_formula)
        
        self.logger.debug(f"Dependências inversas: {dependencias_inversas}")
        
        # 2. Para cada coluna com fórmula, rastrear suas dependências
        formula_dependencies = {}
        for col_formula, formula_info in formulas.items():
            # Extrair nome da coluna sem prefixo
            col_name = col_formula.split('.')[-1] if '.' in col_formula else col_formula
            
            # Obter todas as colunas que esta fórmula referencia
            mapeamento = formula_info.get('mapeamento', {})
            if isinstance(mapeamento, dict):
                deps = list(mapeamento.values())
                formula_dependencies[col_name] = deps
                self.logger.debug(f"Fórmula {col_name} depende de: {deps}")
        
        # 3. Analisar cada resultado de validação
        impacto_por_resultado = {}
        
        for resultado in validation_results:
            linha_idx = resultado.get('linha_idx', 0)
            coluna = resultado['coluna']
            chave = resultado.get('chave', f'Linha_{linha_idx}')
            divergente = not resultado['resultado']
            
            # Se esta coluna tem fórmula e está divergente
            if coluna in formula_dependencies and divergente:
                # Esta é uma coluna calculada com divergência
                # Precisamos identificar qual(is) coluna(s) base causaram isso
                
                # Por enquanto, vamos assumir que TODAS as dependências
                # podem ser a causa (em um cenário real, precisaríamos
                # validar as colunas base também)
                deps = formula_dependencies[coluna]
                
                # Criar mensagem explicativa
                if deps:
                    msg = f"Possíveis causas: divergências em {', '.join(deps[:3])}"
                    if len(deps) > 3:
                        msg += f" e outras {len(deps)-3} colunas"
                else:
                    msg = ""
                
                resultado_key = f"{chave}_{coluna}_{linha_idx}"
                impacto_por_resultado[resultado_key] = msg
                
                self.logger.info(f"Coluna {coluna} linha {linha_idx}: {msg}")
            else:
                # Coluna sem fórmula ou sem divergência
                resultado_key = f"{chave}_{coluna}_{linha_idx}"
                impacto_por_resultado[resultado_key] = ""
        
        return impacto_por_resultado
    
    def gerar_relatorio_impacto(self, impacto_por_resultado: Dict[str, str], 
                               validation_results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Enriquece os resultados de validação com análise de impacto
        """
        resultados_com_impacto = []
        
        for resultado in validation_results:
            linha_idx = resultado.get('linha_idx', 0)
            coluna = resultado['coluna']
            chave = resultado.get('chave', f'Linha_{linha_idx}')
            
            # Criar cópia do resultado
            resultado_enriquecido = resultado.copy()
            
            # Adicionar análise de impacto
            resultado_key = f"{chave}_{coluna}_{linha_idx}"
            divergencias_msg = impacto_por_resultado.get(resultado_key, "")
            
            resultado_enriquecido['divergencias_na_formula'] = divergencias_msg
            
            resultados_com_impacto.append(resultado_enriquecido)
        
        return resultados_com_impacto
    
    def analisar_cadeia_impacto(self, formulas: Dict[str, Dict[str, Any]]) -> Dict[str, Set[str]]:
        """
        Analisa a cadeia de impacto entre fórmulas
        """
        cadeia_impacto = {}
        
        # Para cada fórmula, verificar suas dependências
        for col_formula, formula_info in formulas.items():
            formula_traduzida = formula_info.get('formula_traduzida', '')
            mapeamento = formula_info.get('mapeamento', {})
            
            if isinstance(mapeamento, dict):
                # Para cada coluna referenciada
                for col_ref in mapeamento.values():
                    if col_ref not in cadeia_impacto:
                        cadeia_impacto[col_ref] = set()
                    cadeia_impacto[col_ref].add(col_formula)
                    
                    # Log para debug
                    self.logger.info(f"  {col_ref} impacta: {col_formula}")
        
        return cadeia_impacto