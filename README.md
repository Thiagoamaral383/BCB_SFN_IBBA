# 🚀 Extrator de Séries do BCB (High Performance)

Este projeto consiste em um script Python desenvolvido para substituir e acelerar o processo de extração de dados do Banco Central do Brasil (BCB), anteriormente realizado via VBA.

---

## 📋 Sobre o Projeto

### O Problema

A atualização das séries dependiam de uma planilha Excel utilizando VBA. O processo completo:

- **Tempo médio:** ~2 horas ⏳
- **Instabilidade:** Travamentos frequentes e falhas de conexão.
- **Risco:** Perda de dados ou corrupção da planilha durante o processo.

### A Solução

O script `extract.py` moderniza essa extração, baixando as séries em "Lotes" (Batch) diretamente da API do Banco Central.

- **Tempo médio:** < 4 minutos ⚡
- **Ganhos:** Redução de **97% no tempo de processamento**.
- **Segurança:** Sem travamentos, com validação de dados automática.

> [!IMPORTANT]
> **Atenção:** Este script **NÃO** substitui a planilha mestre de Inteligência de Negócios (BI). Ele apenas realiza a **extração bruta** dos dados. Cálculos complexos (deflação, crescimento real, etc.) continuam sendo feitos no Excel.

---

## 🛠️ Destaques Técnicos

O código foi construído com foco em resiliência ("Enterprise Grade"):

* **🛡️ Robustez e Repescagem:** O script baixa dados em lotes de 10 séries. Se um lote falhar, ele ativa automaticamente o modo de recuperação:
  1. Tenta baixar cada série individualmente.
  2. Se falhar por erro de data, baixa o histórico completo e filtra localmente (útil para séries novas).
* **💾 Backup Automático:** Antes de salvar os novos dados, o script cria automaticamente uma cópia de segurança (`Resultado_BCB_BACKUP.xlsx`) se o arquivo de destino já existir.
* **⚙️ Configuração Centralizada:** Nenhuma alteração de código é necessária para adicionar ou remover séries. Tudo é controlado pelo arquivo `input_series.xlsx`.
* **📂 Portabilidade:** Utiliza caminhos relativos ao diretório de execução. Funciona em qualquer pasta ou máquina sem ajustes.

---

## 🚀 Como Usar

### 1. Pré-requisitos

Certifique-se de ter o Python 3.x instalado. Instale as dependências do projeto:

```bash
pip install -r requirements.txt
```

### 2. Configuração (Opcional)

Se precisar adicionar novas séries, edite o arquivo `input_series.xlsx`. Ele deve conter as colunas:

* **Codigo:** Código da série no SGS/BCB.
* **Coluna:** Coluna de destino no Excel (ex: B, C, AA).
* **Aba:** Nome da aba onde o dado será salvo.

### 3. Execução

Execute o script via terminal na pasta do projeto:

```bash
python extract.py
```

Acompanhe o progresso no terminal. O script mostrará o tempo de execução e status de cada lote.

### 4. Atualização do Dashboard

1. Ao final, abra o arquivo gerado `Resultado_BCB.xlsx`.
2. Copie os dados das abas geradas.
3. Cole na sua planilha mestre de indicadores.
4. Seu Dashboard está atualizado! ✅

---

## 📂 Estrutura de Arquivos

| Arquivo                       | Função                                                                           |
| :---------------------------- | :--------------------------------------------------------------------------------- |
| `extract.py`                | 🐍 Script principal da aplicação. Toda a lógica está aqui.                     |
| `input_series.xlsx`         | ⚙️**Configuração:** Lista de séries a serem baixadas e onde salvá-las. |
| `Resultado_BCB.xlsx`        | 📊**Output:** Arquivo final gerado com os dados atualizados.                 |
| `Resultado_BCB_BACKUP.xlsx` | 🛡️**Segurança:** Backup da execução anterior (gerado automaticamente).  |
| `requirements.txt`          | 📦 Lista de bibliotecas Python necessárias.                                       |

---

## 📞 Suporte

Em caso de erros críticos ("FALHA DEFINITIVA"), verifique:

1. Se o site do BCB (SGS) está no ar.
2. Se o código da série não foi descontinuado.
3. Se o arquivo Excel não está aberto por outro usuário (o que bloqueia a gravação).
