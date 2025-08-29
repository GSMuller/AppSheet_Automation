# Painel Apollo - Automação de Validação de Documentos

## Descrição

Este projeto é voltado para automação de validação de documentos e gerenciamento de cadastramentos administrativos via Google Sheets. O painel foi desenvolvido para facilitar o controle e validação de preenchimento de dados em bancos administrativos, especialmente para equipes de controladoria.

A solução permite:
- Gerenciar e validar documentos de cadastro de forma automatizada.
- Integrar com plataformas administrativas internas.
- Automatizar a verificação de status e preenchimento de colunas específicas.
- Utilizar menus personalizados para operações rápidas (atualizar, guardar, reorganizar dados).

## Funcionalidades

- **Automação de validação:** Verifica automaticamente o preenchimento de documentos e status, aplicando cores e mensagens conforme a situação.
- **Menu personalizado:** Operações de cadastro e documentação acessíveis diretamente pelo menu do Google Sheets.
- **Cache de dados:** Armazena informações temporárias para facilitar reorganização e atualização dos dados.
- **Validação periódica:** Função que pode ser agendada para rodar automaticamente e validar documentos pendentes.
- **Configuração dinâmica:** Adiciona colunas e validações conforme a necessidade do fluxo administrativo.

## Estrutura

- **Cadastro:** Gerenciamento dos dados cadastrais, atualização via integração com planilha Apollo, cache e reorganização.
- **Documentação:** Validação dos documentos, atualização, cache e reorganização dos dados de documentação.
- **Configurações:** Constantes para nomes de abas, cores, status e mapeamento de colunas.
- **Funções auxiliares:** Normalização de textos, mensagens de toast, processamento de bônus/documentos, formatação de células.

## Como usar

1. **Instale o script** no Google Sheets desejado.
2. **Abra a planilha** e utilize o menu "⚙️ Operações" para acessar as funções de cadastro e documentação.
3. **Configure os gatilhos** para rodar funções automáticas, como validação periódica.
4. **Personalize as constantes** conforme a estrutura da sua planilha e fluxo administrativo.

## Créditos

- **Author:** Felipe Marcondes
- **Co-Author:** Giovanni Muller

