# **Gestor de Obras - README**

Bem-vindo ao **Gestor de Obras**, uma solução prática e eficiente para gerenciar seus clientes, funcionários e obras diretamente no Microsoft Excel! Este sistema foi desenvolvido para facilitar o controle de registros, automatizar tarefas rotineiras e melhorar a organização do seu dia a dia.

---

## **📋 Funcionalidades Principais**

### **1. Gestão de Clientes**
Gerencie informações detalhadas de seus clientes:
- **Cadastrar**: Adicione novos clientes com CPF, Nome, Telefone e Endereço.
- **Editar**: Atualize os dados de clientes existentes de forma simples e rápida.
- **Pesquisar**: Encontre informações pelo CPF do cliente.
- **Excluir**: Remova clientes de sua base com segurança, garantindo confirmação prévia.

---

### **2. Gestão de Funcionários**
Organize a equipe envolvida nas obras:
- **Cadastrar**: Adicione os dados dos funcionários, como CPF, Nome, Telefone, Cargo e Data de Nascimento.
- **Editar**: Atualize informações diretamente no formulário.
- **Pesquisar**: Acesse registros de funcionários pelo CPF.
- **Excluir**: Exclua funcionários com apenas um clique.

---

### **3. Gestão de Obras**
Controle detalhado de todas as obras ativas e concluídas:
- **Cadastrar**: Registre informações como ID da Obra, CPF do Cliente, CPF do Funcionário, Endereço, Bairro, Cidade, Prazo, Orçamento e Lucro.
- **Editar**: Ajuste os detalhes das obras sempre que necessário.
- **Pesquisar**: Encontre obras pelo ID e visualize os detalhes completos.
- **Excluir**: Remova obras de maneira simples e direta.

---

## **📂 Estrutura do Sistema**

### **Planilhas**
O sistema é organizado em três planilhas principais:
- **Clientes**: Armazena todos os dados dos clientes cadastrados.
  - **Colunas**: CPF, Nome, Telefone, Endereço.
- **Funcionários**: Contém os registros da equipe.
  - **Colunas**: CPF, Nome, Telefone, Cargo, Data de Nascimento.
- **Obras**: Mantém o histórico das obras cadastradas.
  - **Colunas**: ID da Obra, CPF do Cliente, CPF do Funcionário, Endereço, Bairro, Cidade, Prazo, Orçamento, Lucro.

### **Formulários**
O sistema inclui formulários personalizados para facilitar a interação com os dados:
- **Formulário de Clientes**: Interface intuitiva para gerenciar os dados dos clientes.
- **Formulário de Funcionários**: Registro detalhado da equipe de trabalho.
- **Formulário de Obras**: Gestão completa das informações das obras.

---

## **⚙️ Como Usar**

### **Cadastro de Registros**
1. Abra o formulário correspondente (Clientes, Funcionários ou Obras).
2. Preencha os campos obrigatórios.
3. Clique no botão **Salvar** para registrar os dados.

### **Pesquisa de Registros**
1. Insira o CPF (para Clientes e Funcionários) ou o ID da Obra (para Obras).
2. Clique no botão **Pesquisar** para localizar as informações.

### **Edição de Registros**
1. Realize uma busca pelo CPF ou ID da Obra.
2. Atualize os campos necessários no formulário.
3. Clique no botão **Editar** para salvar as alterações.

### **Exclusão de Registros**
1. Insira o CPF ou ID da Obra no campo de busca.
2. Clique no botão **Excluir** e confirme a operação.

---

## **💻 Requisitos do Sistema**
- Microsoft Excel com suporte a macros habilitado.
- Habilitação para salvar arquivos no formato `.xlsm`.

---

## **🎨 Design Intuitivo**
O **Gestor de Obras** foi projetado para ser funcional e visualmente agradável:
- **Formulários minimalistas**: Priorizam a usabilidade.
- **Mensagens de feedback claras**: Notificam o usuário sobre ações concluídas ou erros.
- **Labels dinâmicos**: Informam sobre o status da operação (como "CPF não encontrado" ou "Dados salvos com sucesso").

---

## **🚀 Instruções de Instalação**
1. Baixe o arquivo `NEQ5_Gestor_Obras.xlsm`.
2. Abra o arquivo no Excel.
3. Habilite as macros ao abrir o arquivo.
4. Navegue pelas planilhas ou use os formulários para começar.

---

## **📌 Convenções de Código**
- Campos seguem a convenção `txt` (para TextBox), `cbb` (para ComboBox) e `lbl` (para Labels):
  - Exemplos: `txtCpfCliente`, `cbbCargo`, `lblPrompt`.
- Botões chamam sub-rotinas específicas:
  - Exemplos: `SalvarClientes`, `EditarFuncionarios`, `ExcluirObra`.

---

## **📖 Documentação Interna**
O código VBA foi amplamente comentado para facilitar a manutenção e personalização:
- **Cada linha explicada**: Com detalhes sobre a funcionalidade.
- **Tratamento de erros**: Inclui verificações para prevenir falhas, como campos vazios ou duplicação de registros.

---

## **🤝 Contribua com o Projeto**
1. Faça um fork do repositório.
2. Realize melhorias no código VBA ou na interface.
3. Documente suas mudanças.
4. Envie um pull request com uma descrição detalhada.

---

Agradecemos por escolher o **Gestor de Obras**! 🚧****
