"""
Módulo responsável por carregar e validar a configuração JSON
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any, List


class ConfigLoader:
    """Carregador e validador de configuração"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.required_fields = {
            'arquivo_fonte_1': ['caminho', 'coluna_chave'],
            'arquivo_fonte_2': ['caminho', 'coluna_chave'],
            'arquivo_formulas': ['caminho', 'colunas_para_carregar', 'mapeamento_formulas'],
            'colunas_para_validar': [],
            'regras_de_tolerancia': [],
            'arquivo_saida': ['caminho']
        }
    
    def load(self, config_path: str) -> Dict[str, Any]:
        """
        Carrega e valida o arquivo de configuração
        
        Args:
            config_path: Caminho para o arquivo config.json
            
        Returns:
            Dict com a configuração validada
            
        Raises:
            FileNotFoundError: Se o arquivo não existir
            ValueError: Se a configuração for inválida
        """
        config_file = Path(config_path)
        
        if not config_file.exists():
            raise FileNotFoundError(f"Arquivo de configuração não encontrado: {config_path}")
        
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(f"Erro ao parsear JSON: {str(e)}")
        
        # Validar configuração
        self._validate_config(config)
        
        # Expandir caminhos relativos
        config = self._expand_paths(config, config_file.parent)
        
        # Aplicar valores padrão
        config = self._apply_defaults(config)
        
        self.logger.info("Configuração carregada e validada com sucesso")
        return config
    
    def _validate_config(self, config: Dict[str, Any]) -> None:
        """Valida se todos os campos obrigatórios estão presentes"""
        # Verificar se é configuração multi-abas ou auto-discovery
        is_multi_sheet = (config.get('arquivo_formulas', {}).get('tipo') == 'multi_abas' or 
                         config.get('tipo_validacao') == 'multi_abas')
        is_auto_discovery = (config.get('arquivo_formulas', {}).get('tipo') == 'auto_discovery' or
                            config.get('tipo_validacao') == 'auto_discovery')
        
        for field, subfields in self.required_fields.items():
            if field not in config:
                raise ValueError(f"Campo obrigatório ausente: {field}")
            
            if subfields and isinstance(config[field], dict):
                for subfield in subfields:
                    # Para configuração multi-abas ou auto-discovery, pular validação de colunas_para_carregar e mapeamento_formulas
                    if (is_multi_sheet or is_auto_discovery) and field == 'arquivo_formulas' and subfield in ['colunas_para_carregar', 'mapeamento_formulas']:
                        continue
                        
                    if subfield not in config[field]:
                        raise ValueError(f"Subcampo obrigatório ausente: {field}.{subfield}")
        
        # Validações específicas
        if not isinstance(config['colunas_para_validar'], list):
            raise ValueError("'colunas_para_validar' deve ser uma lista")
        
        if not config['colunas_para_validar']:
            raise ValueError("'colunas_para_validar' não pode estar vazia")
        
        # Validar mapeamento de fórmulas para configuração padrão
        if not is_multi_sheet and 'mapeamento_formulas' in config['arquivo_formulas']:
            for mapping in config['arquivo_formulas']['mapeamento_formulas']:
                if 'header_resultado' not in mapping:
                    raise ValueError("'header_resultado' ausente no mapeamento de fórmulas")
                if 'header_traducao' not in mapping:
                    raise ValueError("'header_traducao' ausente no mapeamento de fórmulas")
        
        # Validar configuração multi-abas
        if is_multi_sheet and not is_auto_discovery:
            if 'abas' not in config['arquivo_formulas']:
                raise ValueError("'abas' ausente na configuração multi_abas")
            
            for aba_nome, aba_config in config['arquivo_formulas']['abas'].items():
                if 'mapeamento_formulas' in aba_config:
                    for mapping in aba_config['mapeamento_formulas']:
                        if 'header_resultado' not in mapping:
                            raise ValueError(f"'header_resultado' ausente no mapeamento de fórmulas da aba {aba_nome}")
                        if 'header_traducao' not in mapping:
                            raise ValueError(f"'header_traducao' ausente no mapeamento de fórmulas da aba {aba_nome}")
    
    def _expand_paths(self, config: Dict[str, Any], base_path: Path) -> Dict[str, Any]:
        """Converte caminhos relativos em absolutos"""
        def expand_path(path_str: str) -> str:
            path = Path(path_str)
            if not path.is_absolute():
                return str(base_path / path)
            return path_str
        
        # Expandir caminhos dos arquivos
        config['arquivo_fonte_1']['caminho'] = expand_path(config['arquivo_fonte_1']['caminho'])
        config['arquivo_fonte_2']['caminho'] = expand_path(config['arquivo_fonte_2']['caminho'])
        config['arquivo_formulas']['caminho'] = expand_path(config['arquivo_formulas']['caminho'])
        config['arquivo_saida']['caminho'] = expand_path(config['arquivo_saida']['caminho'])
        
        return config
    
    def _apply_defaults(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """Aplica valores padrão quando necessário"""
        # Adicionar regra padrão se não existir
        if 'default' not in config['regras_de_tolerancia']:
            config['regras_de_tolerancia']['default'] = {'tipo': 'exata'}
        
        # Garantir que aba_planilha seja None para CSV
        for arquivo in ['arquivo_fonte_1', 'arquivo_fonte_2']:
            if config[arquivo]['caminho'].lower().endswith('.csv'):
                config[arquivo]['aba_planilha'] = None
        
        return config