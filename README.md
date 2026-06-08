# Sistema de Descarga Operacional — Demo

Projeto demonstrativo em Flask para registro de turnos, descargas, paradas operacionais e visualização em dashboard.

## Conteúdo

- `server.py`: servidor Flask e APIs do sistema.
- `static/relatorio.html`: tela de lançamento do relatório diário.
- `static/dashboard.html`: dashboard de indicadores operacionais.
- `data/relatorios.xlsx`: base demonstrativa preenchida com dados genéricos.
- `scripts/importar_historico.py`: utilitário para importação de histórico.

## Como executar

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python server.py
```

Depois acesse:

- Relatório: `http://localhost:5050`
- Dashboard: `http://localhost:5050/dashboard`

## Base de dados

Por padrão, o sistema usa `data/relatorios.xlsx` dentro do próprio projeto.

Para usar outra pasta de dados, defina a variável de ambiente `APP_DATA` antes de iniciar o servidor.

## Observação

Os dados incluídos são fictícios e foram anonimizados para uso público em repositório Git.
