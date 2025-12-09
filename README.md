# Automação de Classificação e Processamento de Telefones – CPJ + SQL Server

Este projeto automatiza o fluxo completo de classificação, atualização e envio de números telefônicos dentro do sistema legado CPJ, utilizando Python, Selenium e SQL Server.  

## 🚀 Funcionalidades

- Login automático no sistema Gecobi/CPJ  
- Navegação inteligente entre múltiplos iframes  
- Consulta direta ao BD_TELEFONES (SQL Server)  
- Separação automática por classificação:
  - HOT
  - ALTA
  - MEDIA
  - PEQUENA
  - IMPRODUTIVO
- Preenchimento automático do formulário CPJ  
- Seleção do status correspondente no combo do sistema  
- Colagem ultrarrápida de números via JavaScript  
- Execução automatizada da ação “Executar >>”  
- Ciclo completo repetido para cada classificação

## 🛠 Tecnologias Utilizadas

- Python 3  
- Selenium WebDriver  
- PyODBC  
- SQL Server  
- Manipulação avançada de iframes e XPaths  

## 📌 Pré-requisitos

- Chrome + ChromeDriver compatível  
- Driver ODBC SQL Server (17 ou 18)  
- Arquivo de credenciais contendo:
  ```
  CPJ_USER=
  CPJ_PASS=
  BD_TELEFONES_USER=
  BD_TELEFONES_PASS=
  ```

## ▶ Como Executar

1. Clone o repositório  
2. Instale as dependências:  
   ```
   pip install selenium pyodbc
   ```
3. Ajuste o caminho do arquivo de credenciais, se necessário  
4. Execute o script principal:
   ```
   python main.py
   ```

## 📂 Estrutura do Projeto

- **main.py** – script principal da automação  
- Funções de busca recursiva para localizar elementos em qualquer frame  
- Rotinas de classificação e montagem dos blocos de números  
- Rotina de envio para o CPJ  

## 🎯 Objetivo

Automatizar totalmente o processo diário de atualização de status de telefones, eliminando trabalho manual dos operadores, reduzindo erros e garantindo máxima eficiência operacional.

## 📜 Licença

Este projeto pode ser utilizado para fins internos e de automação operacional.  
