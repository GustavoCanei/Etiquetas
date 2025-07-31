# Gerador de Etiquetas

Este projeto é um aplicativo em Python para gerar etiquetas em PDF, com interface gráfica baseada em Tkinter.

## Requisitos

- Python 3.8 ou superior (recomendado instalar do site oficial: https://www.python.org/downloads/)
- Sistema operacional Windows (recomendado)

## Instalação dos pacotes
Abra o PowerShell na pasta do projeto e execute:

```
pip install pandas pillow reportlab python-barcode
```

## Como usar
1. Execute o arquivo `Gerador.py`:
   ```
   python Gerador.py
   ```
2. Preencha os campos na interface.
3. Clique em "Gerar PDF" para salvar as etiquetas.

## Observações
- O arquivo `clients.json` pode ser usado para mapear clientes a logos (opcional).
- Para importar dados de etiquetas, use a função "Buscar Excel" e selecione um arquivo `.xlsx` ou `.xls`.

## Suporte
Em caso de dúvidas ou problemas, entre em contato com o desenvolvedor.
