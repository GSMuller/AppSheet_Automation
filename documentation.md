# :foguete: Painel de Operações Apollo - Documentação
**Author:** Felipe Marcondes
**Co-Author:** Giovanni Muller
**Versão Final:** 1.0
---
## :livro: Introdução
Nós desenvolvemos este código com o objetivo de criar o **Painel de Operações Apollo**, uma automação que integra o **Google Sheets** ao **AppSheet**, trazendo **eficiência, padronização e performance** para o time.
O grande diferencial é a **varredura automática**: como utilizamos o AppSheet para subir informações em tempo real, precisamos garantir que a planilha esteja sempre **consistente**. Com esse código, cada atualização já é processada de forma automática, sem retrabalho manual.
---
## :chave_inglesa: Estrutura do Código
Organizamos o código em blocos bem definidos para facilitar a manutenção e o entendimento:
### 1. Constantes de Configuração
Definimos:
- Nomes das abas principais (**Cadastro**, **Documentação**, **Cache** e **Base Apollo**).
- Paleta de cores para sinalização visual.
- Status possíveis para documentação e validação.
- Mapeamento das colunas de documentos.
:apontando_para_a_direita: Isso facilita futuras alterações, centralizando tudo em um só lugar.
---
### 2. Função `onOpen` - Menu Customizado
Ao abrir a planilha, criamos automaticamente o menu **:engrenagem: Operações** com submenus de **Cadastro** e **Documentação**.
Esse menu permite:
- **:sentido_anti_horário: Atualizar Apollo**
- **:disquete: Guardar Informações**
- **:pino: Reorganizar Dados**
Assim, o time pode rodar funções diretamente pelo Sheets, sem acessar o Apps Script.
---
### 3. Configuração de Coluna de Documentação
Garantimos que a aba **DOCUMENTAÇÃO** tenha a coluna "Documentação" criada e configurada com uma lista suspensa de opções padronizadas.
Isso impede erros de digitação e assegura que os status estejam sempre dentro do esperado.
---
### 4. Eventos (`onEdit`)
O evento `onEdit` é o **coração da automação**.
Sempre que uma linha na aba de **DOCUMENTAÇÃO** é editada, o sistema reage:
- Alteração no campo **Bônus** → processamos quais documentos são obrigatórios.
- Alteração no campo **Status** → ajustamos automaticamente todas as colunas de documentação daquela linha.
:apontando_para_a_direita: Assim, a planilha se mantém coerente com as regras de negócio sem esforço manual.
---
### 5. Validação Periódica
Implementamos também a função `validarDocumentacaoPeriodicamente()`, que realiza uma varredura geral.
Ela garante que nenhuma linha fique desatualizada, mesmo em casos onde o AppSheet realiza edições em massa.
---
### 6. Funções Auxiliares
- **mostrarInfo** → exibe créditos e versão.
- **normalizeText** → padroniza textos (sem acentos, tudo minúsculo).
- **toastMsg** → mostra alertas rápidos no topo da planilha.
São pequenos recursos que aumentam a usabilidade.
---
### 7. Processamento de Documentação
Esta é a parte mais importante do código.
Aqui definimos:
- Quais documentos são obrigatórios de acordo com o tipo de **Bônus**.
- Como preencher automaticamente as colunas de status (**Validado**, **Pendente**, **Não se Aplica**).
- Como aplicar **cores e marcações visuais** para facilitar a análise do time.
Exemplo:
- Para clientes em **Trade-In**, exigimos CRV, CNH e outros documentos específicos.
- Para **Venda Direta**, já sinalizamos que não há documentos obrigatórios extras.
---
## :dardo_no_alvo: Benefícios
- **Automação total** → elimina revisões manuais.
- **Padronização** → todos os status seguem a mesma regra, sem exceções.
- **Performance** → integração com AppSheet em tempo real sem travar a planilha.
- **Escalabilidade** → fácil adaptação caso surjam novos bônus ou documentos.
---
## :anotações: Conclusão
Este código foi criado para **organizar e automatizar** o fluxo operacional.
Nós pensamos nele como um **assistente invisível**, cuidando da parte repetitiva e garantindo que as regras sejam aplicadas corretamente.
Assim, nossa equipe pode se concentrar no que realmente importa: **atender bem os clientes e fechar negócios**, sem desperdício de tempo com verificações manuais.
---
> :mão_escrevendo: Documentação escrita por **Giovanni Muller**, em parceria com **Felipe Marcondes**.