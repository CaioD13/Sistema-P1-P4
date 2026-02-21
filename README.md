# 🛡️ SGI - Sistema de Gestão Integrada (37º BPM/M)

O **SGI (Sistema de Gestão Integrada)** é uma aplicação web desenvolvida em Python (Flask) projetada para modernizar e automatizar as rotinas administrativas de unidades da Polícia Militar, com foco nas seções de Pessoal (P1) e Logística/Finanças (P4).

O sistema substitui o controle manual por planilhas automatizadas, gerencia o efetivo, calcula frequências financeiras e rastreia o fluxo de documentos de forma digital.

## 🚀 Funcionalidades Principais

### 👥 Controle de Efetivo (Módulo P1)
*   **Cadastro Completo:** Registro de Policiais Militares (Posto/Graduação, RE, Cia, Seção, Função).
*   **Gestão de Escalas:** Configuração de regimes de trabalho (12x36 pares/ímpares, Expediente, etc.).
*   **Controle de Saúde:** Monitoramento automático de restrições médicas com prazo de validade.
*   **Edição Individual:** Calendário interativo para alteração pontual de dias de serviço, férias (F), licenças (LP) ou faltas (FS).

### 📊 Automação de Escalas e Relatórios
*   **Escala Semanal Visuais:** Geração automática de arquivo Excel (.xlsx) formatado em A4 Paisagem, agrupado por blocos de horários e Companhias, pronto para afixação em quadro de avisos.
*   **Relatório de Efetivo:** Extração em PDF para conferência rápida de força de trabalho.

### 💰 Controle de Frequência e Finanças (Módulo P4)
*   **Cálculo Automático de Diárias:** O sistema cruza o regime de trabalho do policial com o calendário do mês, preenchendo automaticamente a planilha de frequência.
*   **Regras de Negócio Implementadas:**
    *   `0`: Turnos < 8h (Meio Período).
    *   `1`: Turnos de 8h a 12h (Expediente).
    *   `2`: Turnos de 12h a 18h (Plantão 12h).
    *   `3`: Turnos > 18h (Dobra/24h).
*   **Cálculo de Bônus:** Coluna automatizada que contabiliza pontuação baseada em dias de trabalho e afastamentos remunerados (ignora faltas).

### 🗂️ Protocolo Digital de Documentos
*   **Numeração Única:** Geração automática de protocolo sequencial anual (Ex: `2026/0001`).
*   **Fluxo Completo:** 
    *   `Nova Entrada`: Recebimento de documentos externos.
    *   `Saída Direta`: Expedição de documentos criados na própria seção.
    *   `Trâmite`: Registro de movimentação (quem retirou e quem despachou).
    *   `Arquivamento`: Encerramento do processo com registro do local físico (gaveta/pasta).

---

## 🛠️ Tecnologias Utilizadas

*   **Backend:** Python 3.x, Flask, Flask-SQLAlchemy (ORM), Flask-Login (Autenticação).
*   **Banco de Dados:** SQLite3.
*   **Frontend:** HTML5, CSS3, JavaScript, Bootstrap 5, FontAwesome.
*   **Geração de Relatórios:** XlsxWriter (para formatação avançada de planilhas Excel).

---

## ⚙️ Como Instalar e Rodar o Projeto

### Pré-requisitos
*   [Python 3.x](https://www.python.org/downloads/) instalado na máquina.
*   Gerenciador de pacotes `pip`.

### Passo a Passo

1. **Clone o repositório:**
```bash
   git clone https://github.com/CaioD13/Sistema-P1-P4.git
   cd Sistema-P1-P4
