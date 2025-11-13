#!/usr/bin/env python3
"""
json-to-excel-advanced - Convert complex nested JSON to Excel with multiple sheets

Author: Matheus Nobre Josa
License: MIT
Repository: https://github.com/matheusnorjosa/json-to-excel-advanced
"""

import json
import pandas as pd
import argparse
import sys
from typing import Any, Dict, List, Optional, Union
from datetime import datetime
from collections import Counter
from pathlib import Path


class JSONToExcelConverter:
    """
    Conversor de JSON complexo para Excel com múltiplas abas

    Esta classe processa arquivos JSON aninhados (incluindo exports do MongoDB)
    e gera planilhas Excel organizadas com dados principais, itens detalhados,
    dados brutos e dados normalizados.
    """

    def __init__(self, json_file: str, config: Optional[Dict[str, Any]] = None) -> None:
        """
        Inicializa o conversor

        Args:
            json_file: Caminho para o arquivo JSON de entrada
            config: Dicionário de configurações opcionais
        """
        # Armazena o caminho do arquivo JSON
        self.json_file: str = json_file

        # Configuração padrão ou fornecida pelo usuário
        self.config: Dict[str, Any] = config or {}

        # Dados carregados do JSON (inicialmente None)
        self.data: Optional[List[Dict[str, Any]]] = None

        # Configurações padrão para chaves do JSON
        # Chave que contém itens aninhados (ex: questões, pedidos, etc)
        self.nested_items_key: str = self.config.get('nested_items_key', 'questoes')

        # Chave que contém dados de auditoria com nomes legíveis
        self.audit_key: str = self.config.get('audit_key', 'auditoria')

        # Lista de campos que contêm IDs do MongoDB
        self.id_fields: List[str] = self.config.get('id_fields', [
            '_id', 'aluno', 'turma', 'provaAcompanhamento',
            'acompanhamento', 'municipio', 'corrigoPor'
        ])

    def load_json(self) -> bool:
        """
        Carrega o arquivo JSON

        Returns:
            True se bem-sucedido, False caso contrário
        """
        print(f"📂 Lendo arquivo JSON: {self.json_file}")

        try:
            # Abre e lê o arquivo JSON com encoding UTF-8
            with open(self.json_file, 'r', encoding='utf-8') as f:
                self.data = json.load(f)

            # Verifica se o arquivo está vazio
            if self.data is None:
                print("❌ Erro: Arquivo JSON vazio")
                return False

            # Mostra quantos registros foram carregados
            print(f"✓ Total de registros: {len(self.data)}")
            return True

        except FileNotFoundError:
            # Arquivo não encontrado
            print(f"❌ Erro: Arquivo não encontrado: {self.json_file}")
            return False

        except json.JSONDecodeError as e:
            # JSON inválido ou mal formatado
            print(f"❌ Erro ao decodificar JSON: {e}")
            return False

        except Exception as e:
            # Qualquer outro erro
            print(f"❌ Erro ao ler JSON: {e}")
            return False

    @staticmethod
    def get_oid(obj: Any) -> str:
        """
        Extrai o $oid de um ObjectId do MongoDB

        MongoDB armazena IDs como: {"$oid": "507f1f77bcf86cd799439011"}
        Esta função extrai apenas a string do ID

        Args:
            obj: Objeto que pode conter $oid

        Returns:
            String do ObjectId ou string vazia
        """
        # Verifica se é um dicionário com a chave $oid
        if obj and isinstance(obj, dict):
            return obj.get('$oid', '')
        return ''

    @staticmethod
    def convert_date(date_obj: Any) -> Optional[datetime]:
        """
        Converte data do MongoDB ($date) para datetime do Python

        MongoDB armazena datas como: {"$date": "2025-01-15T10:30:00.000Z"}
        Esta função converte para datetime sem timezone (compatível com Excel)

        Args:
            date_obj: Objeto que pode conter $date

        Returns:
            Objeto datetime ou None
        """
        # Verifica se é um dicionário com a chave $date
        if date_obj and isinstance(date_obj, dict) and '$date' in date_obj:
            try:
                # Converte string de data para datetime usando pandas
                dt = pd.to_datetime(date_obj['$date'])

                # Remove timezone para compatibilidade com Excel
                # Excel não suporta datetime com timezone
                if hasattr(dt, 'tz_localize'):
                    return dt.tz_localize(None)  # type: ignore
                elif hasattr(dt, 'tz_convert'):
                    return dt.tz_convert(None)  # type: ignore
                return dt  # type: ignore

            except Exception:
                # Se falhar, retorna None
                return None
        return None

    @staticmethod
    def safe_get(obj: Any, *keys: str, default: Any = '') -> Any:
        """
        Acessa chaves aninhadas de forma segura em dicionários

        Exemplo: safe_get(data, 'user', 'address', 'city')
        É equivalente a: data['user']['address']['city']
        Mas não dá erro se alguma chave não existir

        Args:
            obj: Dicionário para acessar
            *keys: Chaves para acessar em ordem
            default: Valor padrão se a chave não for encontrada

        Returns:
            Valor na chave aninhada ou valor padrão
        """
        result = obj

        # Percorre cada chave na sequência
        for key in keys:
            # Se o resultado é None, retorna o padrão
            if result is None:
                return default

            # Se é um dicionário, tenta pegar a próxima chave
            if isinstance(result, dict):
                result = result.get(key)
            else:
                # Se não é dicionário, não pode continuar
                return default

        # Retorna o resultado ou padrão se for None
        return result if result is not None else default

    def process_main_data(self) -> pd.DataFrame:
        """
        Processa os dados principais do JSON

        Extrai campos do nível raiz e da seção de auditoria,
        criando uma linha por registro com todas as informações principais

        Returns:
            DataFrame com dados principais
        """
        # Se não há dados carregados, retorna DataFrame vazio
        if self.data is None:
            return pd.DataFrame()

        print("\n📊 Processando dados principais...")

        # Lista para armazenar cada registro processado
        registros_principais: List[Dict[str, Any]] = []

        # Processa cada registro do JSON
        for registro in self.data:
            # Extrai a seção de auditoria (ou dicionário vazio se não existir)
            auditoria: Dict[str, Any] = registro.get(self.audit_key, {}) or {}

            # Cria um dicionário com todos os dados do registro
            linha: Dict[str, Any] = {
                # === IDs do MongoDB ===
                # Extrai ObjectIds usando a função get_oid
                'resultado_id': self.get_oid(registro.get('_id')),
                'aluno_id': self.get_oid(registro.get('aluno')),
                'turma_id': self.get_oid(registro.get('turma')),
                'prova_acompanhamento_id': self.get_oid(registro.get('provaAcompanhamento')),
                'acompanhamento_id': self.get_oid(registro.get('acompanhamento')),
                'municipio_id': self.get_oid(registro.get('municipio')),
                'corrigido_por_id': self.get_oid(registro.get('corrigoPor')),

                # === Nomes legíveis da auditoria ===
                # Extrai nomes em português da seção de auditoria
                'aluno_nome': self.safe_get(auditoria, 'aluno', 'nome'),
                'municipio_nome': self.safe_get(auditoria, 'municipio', 'nome'),
                'acompanhamento_nome': self.safe_get(auditoria, 'acompanhamento', 'nome'),
                'prova_nome': self.safe_get(auditoria, 'provaAcompanhamento', 'nome'),
                'turma_ano': self.safe_get(auditoria, 'turma', 'ano'),
                'turma_turno': self.safe_get(auditoria, 'turma', 'turno'),

                # === Datas ===
                # Converte datas do MongoDB para datetime
                'data_criacao': self.convert_date(registro.get('createdAt')),
                'data_atualizacao': self.convert_date(registro.get('updatedAt')),

                # === Metadados ===
                # Versão do documento (MongoDB)
                '__v': registro.get('__v', ''),

                # Conta quantos itens aninhados existem
                'total_questoes': len(registro.get(self.nested_items_key, [])),
            }

            # === Análise dos itens aninhados ===
            # Extrai a lista de itens aninhados (questões, pedidos, etc)
            items: List[Dict[str, Any]] = registro.get(self.nested_items_key, [])

            # Contadores para análise estatística
            tipos_contador: Counter = Counter()
            formatos_contador: Counter = Counter()
            total_categorias: int = 0

            # Percorre cada item para fazer estatísticas
            for item in items:
                # Conta tipos de itens
                tipo: str = item.get('questaoTipo', '')
                formato: str = item.get('questaoFormato', '')

                if tipo:
                    tipos_contador[tipo] += 1
                if formato:
                    formatos_contador[formato] += 1

                # Conta categorias escolhidas (se existirem)
                categorias = item.get('categoriasEscolhidas')
                if categorias and isinstance(categorias, list):
                    total_categorias += len(categorias)

            # Adiciona o total de categorias à linha
            linha['total_categorias_escolhidas'] = total_categorias

            # Adiciona a linha processada à lista
            registros_principais.append(linha)

        # Converte a lista de dicionários em DataFrame do pandas
        return pd.DataFrame(registros_principais)

    def process_nested_items(self) -> pd.DataFrame:
        """
        Processa itens aninhados em detalhes

        Cada item aninhado (questão, pedido, etc) vira uma linha separada
        na planilha, facilitando análises granulares

        Returns:
            DataFrame com itens detalhados
        """
        # Se não há dados carregados, retorna DataFrame vazio
        if self.data is None:
            return pd.DataFrame()

        print("📋 Processando itens detalhados...")

        # Lista para armazenar cada item processado
        registros_items: List[Dict[str, Any]] = []

        # Processa cada registro do JSON
        for registro in self.data:
            # Extrai informações básicas do registro
            resultado_id: str = self.get_oid(registro.get('_id'))
            auditoria: Dict[str, Any] = registro.get(self.audit_key, {}) or {}
            aluno_nome: str = self.safe_get(auditoria, 'aluno', 'nome')

            # Extrai a lista de itens aninhados
            items: List[Dict[str, Any]] = registro.get(self.nested_items_key, [])

            # Processa cada item individualmente
            for idx, item in enumerate(items, 1):
                # Cria linha com informações básicas do item
                linha_item: Dict[str, Any] = {
                    'resultado_id': resultado_id,
                    'aluno_nome': aluno_nome,
                    'item_numero': idx,  # Número sequencial do item
                    'item_id': self.get_oid(item.get('questaoId')),
                    'item_formato': item.get('questaoFormato', ''),
                    'item_tipo': item.get('questaoTipo', ''),
                }

                # === Adiciona todos os outros campos dinamicamente ===
                # Percorre todos os campos do item
                for key, value in item.items():
                    # Pula o questaoId (já adicionado acima)
                    if key not in ['questaoId']:
                        # Se o valor é uma lista, adiciona contagem e preview
                        if isinstance(value, list):
                            linha_item[f'{key}_count'] = len(value)
                            # Mostra até 5 primeiros itens da lista
                            linha_item[f'{key}_list'] = ', '.join(map(str, value[:5]))

                        # Se é dicionário, converte para string
                        elif isinstance(value, dict):
                            linha_item[key] = str(value)

                        # Qualquer outro tipo, adiciona diretamente
                        else:
                            linha_item[key] = value

                # Adiciona o item processado à lista
                registros_items.append(linha_item)

        # Converte a lista de dicionários em DataFrame
        return pd.DataFrame(registros_items)

    def create_raw_json_sheet(self) -> pd.DataFrame:
        """
        Cria aba com JSON bruto

        Mantém o JSON original completo para referência e backup,
        garantindo que nenhum dado seja perdido na conversão

        Returns:
            DataFrame com JSON completo de cada registro
        """
        # Se não há dados carregados, retorna DataFrame vazio
        if self.data is None:
            return pd.DataFrame()

        print("💾 Criando aba de dados brutos...")

        # Lista para armazenar JSON de cada registro
        registros_brutos: List[Dict[str, str]] = []

        # Processa cada registro
        for registro in self.data:
            linha_bruta: Dict[str, str] = {
                'resultado_id': self.get_oid(registro.get('_id')),
                # Converte o registro inteiro para JSON formatado
                'json_completo': json.dumps(registro, ensure_ascii=False, indent=2)
            }
            registros_brutos.append(linha_bruta)

        # Converte para DataFrame
        return pd.DataFrame(registros_brutos)

    def create_normalized_data(self) -> pd.DataFrame:
        """
        Cria dados completamente achatados/normalizados

        Usa json_normalize do pandas para achatar todas as estruturas aninhadas,
        transformando cada campo aninhado em uma coluna separada

        Returns:
            DataFrame com dados normalizados
        """
        # Se não há dados carregados, retorna DataFrame vazio
        if self.data is None:
            return pd.DataFrame()

        print("🔧 Normalizando dados (JSON achatado)...")

        try:
            # Usa json_normalize para achatar o JSON
            # max_level=None: achata todos os níveis
            # sep='_': usa underscore para separar níveis (ex: user_address_city)
            df_normalized: pd.DataFrame = pd.json_normalize(self.data, max_level=None, sep='_')

            # Limpa nomes de colunas
            # Remove $ (do MongoDB) e substitui . por _
            df_normalized.columns = [
                str(col).replace('$', '').replace('.', '_')
                for col in df_normalized.columns
            ]

            return df_normalized

        except Exception as e:
            # Se houver erro na normalização, mostra aviso e retorna vazio
            print(f"⚠ Aviso: {e}")
            return pd.DataFrame()

    def convert(self, output_file: Optional[str] = None) -> bool:
        """
        Converte JSON para Excel

        Método principal que orquestra todo o processo de conversão:
        1. Carrega o JSON
        2. Processa todos os dados
        3. Cria o arquivo Excel com múltiplas abas
        4. Ajusta formatação

        Args:
            output_file: Caminho do arquivo Excel de saída (opcional)

        Returns:
            True se bem-sucedido, False caso contrário
        """
        # === 1. CARREGAMENTO ===
        # Carrega o arquivo JSON
        if not self.load_json():
            return False

        # === 2. GERAR NOME DO ARQUIVO DE SAÍDA ===
        # Se não foi fornecido, usa o mesmo nome do JSON com extensão .xlsx
        if not output_file:
            input_path = Path(self.json_file)
            output_file = str(input_path.with_suffix('.xlsx'))

        # === 3. PROCESSAMENTO ===
        # Processa todos os tipos de dados
        print("\n⚙️ Processando dados...")
        df_principal: pd.DataFrame = self.process_main_data()
        df_items: pd.DataFrame = self.process_nested_items()
        df_brutos: pd.DataFrame = self.create_raw_json_sheet()
        df_normalized: pd.DataFrame = self.create_normalized_data()

        # === 4. SALVAMENTO NO EXCEL ===
        print(f"\n💾 Salvando arquivo Excel: {output_file}")

        try:
            # Cria o writer do Excel usando openpyxl
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

                # === ESCREVE CADA ABA ===
                # Aba 1: Dados Principais
                if not df_principal.empty:
                    df_principal.to_excel(
                        writer,
                        sheet_name='1. Dados Principais',
                        index=False  # Não inclui índice do DataFrame
                    )

                # Aba 2: Itens Detalhados
                if not df_items.empty:
                    df_items.to_excel(writer, sheet_name='2. Itens Detalhados', index=False)

                # Aba 3: JSON Bruto (backup completo)
                if not df_brutos.empty:
                    df_brutos.to_excel(writer, sheet_name='3. Dados Brutos JSON', index=False)

                # Aba 4: Dados Normalizados (JSON achatado)
                if not df_normalized.empty:
                    df_normalized.to_excel(writer, sheet_name='4. Dados Normalizados', index=False)

                # === AJUSTA LARGURA DAS COLUNAS ===
                # Melhora a visualização ajustando largura baseada no conteúdo
                for sheet_name in writer.sheets:
                    # Para todas as abas exceto JSON bruto
                    if sheet_name != '3. Dados Brutos JSON':
                        worksheet = writer.sheets[sheet_name]

                        # Percorre cada coluna
                        for column in worksheet.columns:
                            max_length: int = 0
                            column_cells: List[Any] = [cell for cell in column]

                            # Encontra o tamanho máximo de conteúdo na coluna
                            for cell in column_cells:
                                try:
                                    cell_length = len(str(cell.value))
                                    if cell_length > max_length:
                                        max_length = cell_length
                                except Exception:
                                    pass

                            # Ajusta largura (mínimo 2, máximo 50)
                            adjusted_width: int = min(max_length + 2, 50)
                            worksheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

                    else:
                        # Para JSON bruto, usa largura fixa grande
                        worksheet = writer.sheets[sheet_name]
                        worksheet.column_dimensions['B'].width = 100

            # === 5. RESUMO ===
            # Mostra resumo da conversão
            print(f"\n{'='*70}")
            print(f"✓ CONVERSÃO CONCLUÍDA COM SUCESSO!")
            print(f"{'='*70}")
            print(f"\n📄 Arquivo: {output_file}")
            print(f"\n📊 RESUMO:")
            print(f"   • Registros principais: {len(df_principal)}")
            print(f"   • Itens detalhados: {len(df_items)}")
            print(f"   • Abas criadas: {len(writer.sheets)}")
            print(f"\n✓ 100% dos dados preservados!")
            print(f"{'='*70}\n")

            return True

        except Exception as e:
            # Se houver erro ao salvar, mostra mensagem
            print(f"\n❌ Erro ao salvar Excel: {e}")
            return False


def main() -> None:
    """
    Ponto de entrada da CLI (Command Line Interface)

    Configura e processa argumentos da linha de comando,
    permitindo uso do script via terminal
    """
    # === CONFIGURAÇÃO DO PARSER DE ARGUMENTOS ===
    parser = argparse.ArgumentParser(
        description='Convert complex nested JSON to Excel with multiple sheets',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s data.json
  %(prog)s data.json -o output.xlsx
  %(prog)s data.json --nested-key items --audit-key metadata

For more information: https://github.com/matheusnorjosa/json-to-excel-advanced
        """
    )

    # === ARGUMENTOS DA CLI ===
    # Argumento obrigatório: arquivo de entrada
    parser.add_argument('input', help='Input JSON file')

    # Argumento opcional: arquivo de saída
    parser.add_argument('-o', '--output', help='Output Excel file (default: input.xlsx)')

    # Argumento opcional: chave dos itens aninhados
    parser.add_argument(
        '--nested-key',
        default='questoes',
        help='Key for nested items (default: questoes)'
    )

    # Argumento opcional: chave da auditoria
    parser.add_argument(
        '--audit-key',
        default='auditoria',
        help='Key for audit/metadata (default: auditoria)'
    )

    # Mostra versão
    parser.add_argument('--version', action='version', version='%(prog)s 1.0.0')

    # Parse dos argumentos
    args = parser.parse_args()

    # === VALIDAÇÃO ===
    # Verifica se o arquivo de entrada existe
    if not Path(args.input).exists():
        print(f"❌ Erro: Arquivo não encontrado: {args.input}")
        sys.exit(1)

    # === CONVERSÃO ===
    # Cria dicionário de configuração com os argumentos
    config: Dict[str, str] = {
        'nested_items_key': args.nested_key,
        'audit_key': args.audit_key
    }

    # Cria o conversor e executa
    converter = JSONToExcelConverter(args.input, config)
    success: bool = converter.convert(args.output)

    # Retorna código de saída (0 = sucesso, 1 = erro)
    sys.exit(0 if success else 1)


# === PONTO DE ENTRADA ===
# Se o script for executado diretamente (não importado)
if __name__ == '__main__':
    main()
