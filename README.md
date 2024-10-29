# ExWord System v1.0

## Descrição
O **ExWord System** é um sistema em VBA desenvolvido para facilitar a integração entre o Excel e o Word. Ele permite substituir automaticamente campos em arquivos Word (extensões `.docx` e `.xml`), salvando os documentos preenchidos em formato `.pdf` ou `.docx`. 

Este sistema é útil para gerar documentos personalizados a partir de um template com variáveis dinâmicas, automatizando o fluxo de criação de relatórios, contratos, certificados, e outros documentos.

## Funcionalidades

- **Substituição Automática**: Faz a busca e substituição de campos personalizados no arquivo Word.
- **Conversão para PDF**: Converte o arquivo Word finalizado diretamente para PDF.
- **Interface Personalizável**: Interface de UserForm com campos para selecionar o diretório de saída, template do Word, e intervalos de chaves e valores no Excel.
- **Formatação Automática**: Formata o formulário e os controles de entrada de acordo com o design definido.
- **Extração e Compressão**: Manipula arquivos `.docx` e `.xml` diretamente, substituindo os campos no XML quando necessário.
  
## Como Utilizar

1. **Selecionar Diretório de Saída**: Clique em "Selecionar Diretório de Saída" e escolha a pasta onde os documentos gerados serão salvos.
2. **Selecionar Template do Word**: Escolha um arquivo `.docx` ou `.xml` com os campos de template.
3. **Selecionar Campos**: Utilize "Selecione Campos" para indicar os intervalos de "chaves" e "valores" no Excel que substituirão os campos no template do Word.
4. **Exportar Documento**: Clique em "Exportar Documento" para gerar o documento final, que será salvo em formato PDF e `.docx` no diretório selecionado.

## Estrutura do Código

- `ConverterWordParaPDF`: Função para abrir o Word e converter documentos em PDF.
- `SubstituirCamposArquivoWord`: Realiza a busca e substituição de campos no documento Word ou no XML, quando aplicável.
- `ZipUnzipFile`: Função de extração e compressão de arquivos `.zip`, usada para manipular o conteúdo de `.docx`.
- `StringToEntities`: Transforma caracteres especiais em entidades XML para evitar erros de compatibilidade.
- `IniciarDicionario`: Cria o dicionário de chaves e valores a serem substituídos no documento.
- `exportarDados_Click`: Botão para iniciar o processo de exportação.
  
## Requisitos

- Microsoft Excel e Word (Office 2013 ou superior)
- Referências VBA habilitadas:
  - Microsoft Scripting Runtime
  - Microsoft Shell Controls And Automation
  - Microsoft Word 16.0 Object Library

Obs: Embora a habilitação das referências não seja necessária, é recomendado, e após habilitadas altere #Const EarlyBind para True

## Configurações

#Const EarlyBind = True ' Defina como True para Early Binding ou False para Late Binding
