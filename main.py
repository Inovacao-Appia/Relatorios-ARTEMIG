import os
import sys
import datetime
import shutil
import importlib.util
from pathlib import Path

# ====== ESTRUTURA DE PASTAS ======
ESTRUTURA = {
    'A-OPERACAO': [
        'APREENSAO-DE-ANIMAIS',
        'ATENDIMENTO-MECANICO',
        'ATENDIMENTO-PRE-HOSPITALAR',
        'CICLO-INSPECAO',
        'COMBATE-A-INCENDIO',
        'ESTATISTICA-DE-ACIDENTE',
        'ESTATISTICA-DE-TRAFEGO',
        'TELEMETRIA-DE-VEICULOS',
        'TEMPO-DE-ATENDIMENTO'
    ],
    'B-RADARES': [
        'RADARES'
    ],
    'C-SATs': [
        'SATs'
    ],
    'D-DISPONIBILIDADE': [
        'DISPONIBILIDADE-BALANCA-HSWIN',
        'DISPONIBILIDADE-BALANCA-MOVEL',
        'DISPONIBILIDADE-BALANCA-PUNITIVA',
        'DISPONIBILIDADE-BALANCA-SELETIVA',
        'DISPONIBILIDADE-CFTV-OPERACIONAL',
        'DISPONIBILIDADE-DETECTOR-ALTURA',
        'DISPONIBILIDADE-PMV-MOVEL',
        'DISPONIBILIDADE-RADAR',
        'DISPONIBILIDADE-RADIOCOMUNICACAO-LTE',
        'DISPONIBILIDADE-SAT',
        'DISPONIBILIDADE-SITE-INTERNET',
        'DISPONIBILIDADE-TELEFONIA'
    ],
    'E-PESAGEMVEICULAR': [
        'AFERICAO',
        'PROGRAMACAO-DE-FISCALIZACAO-E-MANUTENCAO',
        'RELATORIO-MENSAL-DE-PESAGEM',
        'RELATORIO-MENSAO-DE-PESAGEM-E-AUTUACAO'
    ]
}

def obter_mes_atual():
    """Retorna o mês atual no formato YYYY-MM"""
    hoje = datetime.date.today()
    return f"{hoje.year:04d}-{hoje.month:02d}"

def criar_estrutura_pastas(base_dir: Path, mes_ref: str):
    """Cria estrutura de pastas para o mês"""
    pasta_mes = base_dir / mes_ref
    pasta_mes.mkdir(parents=True, exist_ok=True)
    
    for pasta_principal, subpastas in ESTRUTURA.items():
        for subpasta in subpastas:
            caminho = pasta_mes / pasta_principal / subpasta
            caminho.mkdir(parents=True, exist_ok=True)
    
    return pasta_mes

def importar_modulo(caminho_py: Path):
    """Importa dinamicamente um módulo Python"""
    spec = importlib.util.spec_from_file_location(caminho_py.stem, caminho_py)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo

def executar_relatorios(pasta_destino: Path, mes_ref: str):
    """
    Busca e executa todos os scripts Python como funções
    """
    pasta_python = Path("Python")
    resultados = []
    
    print("\n" + "="*60)
    print("EXECUTANDO RELATÓRIOS")
    print("="*60)
    
    # Percorrer estrutura procurando arquivos .py
    for pasta_principal, subpastas in ESTRUTURA.items():
        for subpasta in subpastas:
            # Caminho esperado do script
            caminho_script = pasta_python / pasta_principal / subpasta / f"{subpasta}.py"
            
            if caminho_script.exists():
                try:
                    # Importar o módulo
                    modulo = importar_modulo(caminho_script)
                    
                    # Verificar se tem a função gerar_relatorio
                    if hasattr(modulo, 'gerar_relatorio'):
                        # Chamar a função
                        arquivo_gerado = modulo.gerar_relatorio(pasta_destino, mes_ref)
                        resultados.append((caminho_script, True, arquivo_gerado))
                    else:
                        print(f"⚠️  {caminho_script.name}: não tem função 'gerar_relatorio'")
                        resultados.append((caminho_script, False, "Função não encontrada"))
                        
                except Exception as e:
                    print(f"❌ Erro em {caminho_script.name}: {e}")
                    resultados.append((caminho_script, False, str(e)))
            else:
                # Script não existe ainda (não implementado)
                pass
    
    return resultados

def criar_zip(pasta_mes: Path):
    """Cria arquivo ZIP com todos os relatórios"""
    nome_zip = str(pasta_mes)
    shutil.make_archive(nome_zip, 'zip', pasta_mes)
    return Path(f"{nome_zip}.zip")

def main():
    """
    Função principal do orquestrador
    
    Uso:
        python main.py              -> usa mês atual
        python main.py 2026-01      -> usa mês específico
    """
    
    # Definir mês
    if len(sys.argv) > 1:
        mes_ref = sys.argv[1]
    else:
        mes_ref = obter_mes_atual()
    
    print("\n" + "="*60)
    print(f"GERADOR DE RELATÓRIOS - MÊS: {mes_ref}")
    print("="*60)
    
    # Criar estrutura
    base_excel = Path("Excel")
    pasta_mes = criar_estrutura_pastas(base_excel, mes_ref)
    print(f"\n📁 Estrutura criada em: {pasta_mes}")
    
    # Executar relatórios
    resultados = executar_relatorios(pasta_mes, mes_ref)
    
    # Resumo
    print("\n" + "="*60)
    print("RESUMO DA EXECUÇÃO")
    print("="*60)
    
    sucesso = sum(1 for _, ok, _ in resultados if ok)
    total = len(resultados)
    
    for script, ok, info in resultados:
        status = "✅" if ok else "❌"
        print(f"{status} {script.stem}")
    
    print(f"\nTotal: {sucesso}/{total} relatórios gerados com sucesso")
    
    # Criar ZIP
    if sucesso > 0:
        print("\n📦 Criando arquivo ZIP...")
        arquivo_zip = criar_zip(pasta_mes)
        print(f"✅ ZIP criado: {arquivo_zip}")
        print(f"   Tamanho: {arquivo_zip.stat().st_size / 1024 / 1024:.2f} MB")
    
    print("\n" + "="*60)
    print("PROCESSO FINALIZADO!")
    print("="*60)

if __name__ == "__main__":
    main()
