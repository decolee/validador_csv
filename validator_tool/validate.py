#!/usr/bin/env python3
"""
Ferramenta de Validação e Análise de Concordância de Dados e Fórmulas
Validador para migração de cálculos Excel para SQL Server
"""

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Dict, Any, List, Tuple, Optional
import pandas as pd

from src.config_loader import ConfigLoader
from src.data_loader import DataLoader
from src.data_aligner import DataAligner
from src.validator import Validator
from src.formula_extractor import FormulaExtractor
from src.report_generator import ReportGenerator
from src.impact_analyzer import ImpactAnalyzer


def setup_logging(verbose: bool = False) -> None:
    """Configura o sistema de logging"""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('validator.log')
        ]
    )


def main():
    """Função principal do validador"""
    parser = argparse.ArgumentParser(
        description='Validador de Dados e Fórmulas - Migração Excel para SQL'
    )
    parser.add_argument(
        'config_file',
        type=str,
        help='Caminho para o arquivo de configuração JSON'
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Ativa modo verbose (debug)'
    )
    
    args = parser.parse_args()
    
    # Setup logging
    setup_logging(args.verbose)
    logger = logging.getLogger(__name__)
    
    try:
        # 1. Carregar configuração
        logger.info(f"Carregando configuração de: {args.config_file}")
        config_loader = ConfigLoader()
        config = config_loader.load(args.config_file)
        
        # 2. Carregar dados dos arquivos
        logger.info("Carregando arquivos de dados...")
        data_loader = DataLoader()
        
        # Arquivo fonte 1
        df1 = data_loader.load_file(
            config['arquivo_fonte_1']['caminho'],
            config['arquivo_fonte_1'].get('aba_planilha')
        )
        logger.info(f"Arquivo 1 carregado: {len(df1)} linhas")
        
        # Arquivo fonte 2
        df2 = data_loader.load_file(
            config['arquivo_fonte_2']['caminho'],
            config['arquivo_fonte_2'].get('aba_planilha')
        )
        logger.info(f"Arquivo 2 carregado: {len(df2)} linhas")
        
        # 3. Alinhar dados usando coluna chave
        logger.info("Alinhando dados usando coluna chave...")
        aligner = DataAligner()
        df1_aligned, df2_aligned, unmatched = aligner.align_dataframes(
            df1,
            df2,
            config['arquivo_fonte_1']['coluna_chave'],
            config['arquivo_fonte_2']['coluna_chave']
        )
        
        if unmatched['df1'] or unmatched['df2']:
            logger.warning(f"Linhas sem correspondência - Arquivo 1: {len(unmatched['df1'])}, "
                         f"Arquivo 2: {len(unmatched['df2'])}")
        
        # 4. Validar dados
        logger.info("Iniciando validação de dados...")
        validator = Validator(config['regras_de_tolerancia'])
        
        # Obter coluna chave para incluir nos resultados
        key_column = config['arquivo_fonte_1']['coluna_chave']
        
        validation_results = validator.validate(
            df1_aligned,
            df2_aligned,
            config['colunas_para_validar'],
            key_column
        )
        
        # 5. Extrair e traduzir fórmulas
        logger.info("Extraindo e traduzindo fórmulas...")
        formula_extractor = FormulaExtractor()
        formulas = formula_extractor.extract_and_translate(
            config['arquivo_formulas'],
            config['arquivo_fonte_1']['coluna_chave']
        )
        
        logger.info(f"Fórmulas extraídas: {list(formulas.keys())}")
        
        # 6. Análise de impacto de divergências
        logger.info("Analisando impacto de divergências nas fórmulas...")
        impact_analyzer = ImpactAnalyzer()
        
        # Analisar dependências e impactos
        impacto_por_resultado = impact_analyzer.analyze_formula_dependencies(
            validation_results,
            formulas
        )
        
        # Enriquecer resultados com análise de impacto
        validation_results_com_impacto = impact_analyzer.gerar_relatorio_impacto(
            impacto_por_resultado,
            validation_results
        )
        
        # Analisar cadeia de impacto
        cadeia_impacto = impact_analyzer.analisar_cadeia_impacto(formulas)
        
        # 7. Preparar informações extras para o relatório
        extra_info = {}
        
        # Dependências
        dependencias = {
            'formulas': formulas,
            'reverse_dependencies': {},
            'statistics': {
                'total_formulas': len(formulas),
                'total_dependencies': sum(len(f.get('depends_on', [])) for f in formulas.values()),
                'cross_sheet_deps': sum(len(f.get('referencias_externas', [])) for f in formulas.values()),
                'most_referenced': [],
                'max_depth': 0
            }
        }
        
        # Calcular dependências reversas
        for col_name, formula_info in formulas.items():
            for dep in formula_info.get('depends_on', []):
                if dep not in dependencias['reverse_dependencies']:
                    dependencias['reverse_dependencies'][dep] = []
                dependencias['reverse_dependencies'][dep].append(col_name)
        
        # Encontrar colunas mais referenciadas
        ref_counts = {}
        for deps in dependencias['reverse_dependencies'].values():
            for dep in deps:
                ref_counts[dep] = ref_counts.get(dep, 0) + 1
        
        if ref_counts:
            sorted_refs = sorted(ref_counts.items(), key=lambda x: x[1], reverse=True)
            dependencias['statistics']['most_referenced'] = [r[0] for r in sorted_refs[:3]]
        
        extra_info['dependencias'] = dependencias
        
        # Alertas
        alertas = []
        
        # Analisar resultados para gerar alertas
        if validation_results:
            df_results = pd.DataFrame(validation_results)
            
            # Alertas por alta taxa de erro
            for col in config['colunas_para_validar']:
                col_results = df_results[df_results['coluna'] == col]
                if len(col_results) > 0:
                    error_rate = len(col_results[col_results['resultado'] == False]) / len(col_results)
                    if error_rate > 0.3:
                        alertas.append({
                            'severidade': 'alta',
                            'tipo': 'taxa_erro_alta',
                            'localizacao': col,
                            'descricao': f'Coluna {col} com {error_rate*100:.1f}% de divergências',
                            'recomendacao': 'Revisar processo de cálculo ou mapeamento desta coluna',
                            'impacto': f'{len(col_results)} registros afetados'
                        })
                    elif error_rate > 0.1:
                        alertas.append({
                            'severidade': 'media',
                            'tipo': 'taxa_erro_media',
                            'localizacao': col,
                            'descricao': f'Coluna {col} com {error_rate*100:.1f}% de divergências',
                            'recomendacao': 'Verificar casos específicos de divergência',
                            'impacto': f'{len(col_results[col_results["resultado"] == False])} registros afetados'
                        })
            
            # Alertas por impacto em cascata
            for col, impactos in cadeia_impacto.items():
                if len(impactos) > 3:
                    alertas.append({
                        'severidade': 'alta',
                        'tipo': 'impacto_cascata',
                        'localizacao': col,
                        'descricao': f'Divergências em {col} afetam {len(impactos)} outras colunas',
                        'recomendacao': 'Priorizar correção desta coluna devido ao alto impacto',
                        'impacto': f'Afeta: {", ".join(list(impactos)[:3])} e outras'
                    })
        
        # Alertas de configuração
        if 'analises_adicionais' in config and 'alertas_personalizados' in config['analises_adicionais']:
            for alerta_config in config['analises_adicionais']['alertas_personalizados']:
                # Aqui seria implementada a lógica específica de cada alerta personalizado
                pass
        
        extra_info['alertas'] = alertas
        
        # Impacto em cascata
        impacto_cascata = {}
        for col_origem, cols_afetadas in cadeia_impacto.items():
            if cols_afetadas:
                # Separar colunas e fórmulas
                colunas = []
                formulas_afetadas = []
                
                for item in cols_afetadas:
                    # Se tem ponto, é uma fórmula (ex: Producao.SUBTOTAL)
                    if '.' in item:
                        formulas_afetadas.append(item)
                    else:
                        colunas.append(item)
                
                impacto_cascata[col_origem] = {
                    'affected': {
                        'columns': colunas,
                        'formulas': formulas_afetadas
                    },
                    'divergent_value': 'Ver detalhes'
                }
        
        extra_info['impacto_cascata'] = impacto_cascata
        extra_info['tipo_validacao'] = config.get('tipo_validacao', 'Padrão')
        
        # 8. Gerar relatório com informações extras
        logger.info("Gerando relatório de validação...")
        report_generator = ReportGenerator()
        report_generator.generate(
            validation_results_com_impacto,
            formulas,
            unmatched,
            config['arquivo_saida']['caminho'],
            extra_info
        )
        
        logger.info(f"Validação concluída! Relatório salvo em: {config['arquivo_saida']['caminho']}")
        
    except Exception as e:
        logger.error(f"Erro durante a validação: {str(e)}", exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()