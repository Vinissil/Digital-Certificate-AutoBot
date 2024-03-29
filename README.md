<h1>Automação de Certificado Digital</h1>

<h1>Digital-Certificate-AutoBot</h1>

> Estado: Projeto Aberto

> Status: Open Project

### Projeto Aberto em Python - Projeto de Automação 
### Open Python Project - Automation Project

### Desenvolvedor, Colaborador - Creditos:
|Nome|Canal do Youtube|Arquivos|
| -------- | -------- | -------- |
| Jhonatan de Souza | [Dev Aprender](https://www.youtube.com/watch?v=VwYqakOB4ow&ab_channel=DevAprender%7CJhonatandeSouza)| [Arquivos da aula](https://drive.google.com/drive/folders/184jpLrEPsmQ-vI2n0LcZSiCmu7_86iqB)

### Developer, Collaborator - Credits:
|Name|Youtube Channel|Archives|
| -------- | -------- | -------- |
| Jhonatan de Souza | [Dev Aprender](https://www.youtube.com/watch?v=VwYqakOB4ow&ab_channel=DevAprender%7CJhonatandeSouza)| [Class Archives](https://drive.google.com/drive/folders/184jpLrEPsmQ-vI2n0LcZSiCmu7_86iqB)

### Projeto executado para fins de estudo apenas - pelo estudante :
### Project executed only for study purposes - by the student :

|Nome|Curso|Linkedin|
| -------- | -------- | -------- |
| Vinicius | [Engenharia de Software - FIAP] | [Perfil](https://www.linkedin.com/in/vinicius-souza-e-silva-1333a5281/)


### Conteúdo 

- Todos os conteudos são encontrados através do link do canal ou diretamente no link do [Drive](https://drive.google.com/drive/folders/184jpLrEPsmQ-vI2n0LcZSiCmu7_86iqB) disponível nesse arquivo README.md

### Content

- All content is found through the channel link or directly in the [Drive](https://drive.google.com/drive/folders/184jpLrEPsmQ-vI2n0LcZSiCmu7_86iqB) link available in this README.md file

### Descrição

- Aplicação ultilizada para automatizar a emissão de certificados de maineira mais rápida de acordo com um layout pré-estabelecido pela instituição.
- O layout conta com os seguintes requisitos exigidos pela insituição, tais como o nome do curso, nome participante, tipo de participação, data do inicio, data do final, carga horária, data da emissão do certificado.
- Ultilizando um arquivo xlsx que seria um documento padrão do excel.

- Quero ver a possibilidade de criar um programa usando o Python para automatizar enviando os dados da planilha para preencher os campos mutáveis no certificado padrão.

- Tipo nome do curso, nome participante, tipo de participação, data do inicio, data do final, carga horária, data da emissão do certificado e as assinaturas do Gestor Geral, do Coordenador e do aluno.

- Pegar dados(texto) -> sobrepor de uma imagem(certificado padrão)

 - Pegar os dados da planilha
    - Tais como nome do curso, nome participante, tipo de participação, data do inicio, data do final, carga horária, data da emissão do certificado

- Tranferir para a imagem do certificado

### Description

- Application used to automate the issuance of main certificates faster according to a layout pre-established by the institution.
- The layout has the following requirements required by the institution, such as the name of the course, participant name, type of participation, start date, end date, workload, date of issuance of the certificate.
- Using an xlsx file that would be a standard Excel document.

- I want to see the possibility of creating a program using Python to automate sending spreadsheet data to populate the changeable fields in the default certificate.

- Like course name, participant name, type of participation, start date, end date, workload, certificate issuance date and the signatures of the General Manager, the Coordinator and the student.

- Get data (text) -> overlay from an image (standard certificate)

- Get the data from the spreadsheet
    - Type course name, participant name, type of participation, start date, end date, workload, certificate issuance date

- Transfer to certificate image

### Layout

[<img src="https://github.com/Vinissil/Digital-Certificate-AutoBot/blob/master/certificado_padrao.jpg" width=60%>]
| :---: |

### Dependências e Libs Instaladas - Dependencies and Installed Libs

- <h3>openpyxl</h3> -
    - import openpyxl
- <h3>Pillow</h3> -
    - from PIL import Image, ImageDraw, ImageFont

### Como rodar a aplicação - Running the application 

- Construir um hambiente virtual com Python 3.
    - Utilizar os comandos : 
        - python -m venv [nome do ambiente virtual](sem os couxetes)(dar enter)
    - ativar o ambiente virtual : 
        - .\[nome do ambiente virtual](sem os couxetes)\Scripts\activate (dar enter)
    - Instalando as Bibliotecas : 
        - pip install pillow openpyxl (dar enter e aguardar o tempo de instalação)

- Trazer os arquivos que serão ultilizados na pasta, para melhorar o armazenamento dos certificados quando emitidos, fazer uma pasta para não ficarem desorganizados.

### Como foi executado o passo a passo



<h4>Português</h4>

<h6>Passo a Passo</h6>

- 1° Pegar os dados da planilha.

- 2º Importar bibliotecas usadas para rodar o codigo openpyxl(import openpyxl) para abrir o arquivo excel e o pillow (from PIL import Image, ImageDraw, ImageFont) para abrir a imagem do certificado.

- 3° Criar as variaveis para abrir a planilha e a aba da planilha.

- 4° Colocar um indice para cada linha da planilha e colocar um liite minimo e maximo para gerar os certificados durante o periodo de teste.

- 5° Criar a célula que contém a info que é preciso para encontrar a linha correta dos requisitos do certificado.

- 6° e 7° Tranferir os dados da planilha para a imagem do certificado. Definindo fonte a ser usada. Definindo fonte a ser usada.

- 8º Crie o caminho para abrir a imagem do certificado.

- 9° Adicione as informações ao camposrespectivos no certificado.

- 10° Para o programa ler datas e carga horária, é preciso mudar a class de número para string.

- 11° Assim é feito com a data de inicio e de termino.

- 12° Criar uma pasta para guardar as imagens geradas pelo bot para melhor organização. 

- 13º Rodar o codigo para gerar o primeiro certificado.

### Qual quer dúvida apenas assista o video.

### How to execute

<h4>English</h4>

<h6>Steps</h6>

- 1° Get the data from the spreadsheet.

- 2º Import libraries used to run the code openpyxl(import openpyxl) to open the excel file and the pillow(from PIL import Image, ImageDraw, ImageFont) to open the certificate image.

- 3° Create the variables to open the spreadsheet and the spreadsheet tab.

- 4° Put an index for each row of the spreadsheet and put a minimum and maximum limit to generate the certificates during the test period.

- 5° Create the cell that contains the info that is needed to find the correct line of the certificate requirements.

- 6° and 7° Transfer the data from the spreadsheet to the certificate image. Defining font to be used.

- 8º Create the path to open the certificate image.

- 9° Adding the information to the respective fields in the certificate.

- 10° For the program to read dates and workload, you need to change the number class to string.

- 11° The same need to be done with the start and end date.

- 12° and  Create a folder to save the images generated by the bot for better organization. 

- 13° Run the code to generate the first certificate.

### If you have any questions, just watch the video. which has English subtitles in automatic translation.


### Liscença - License

                    GNU GENERAL PUBLIC LICENSE
                       Version 3, 29 June 2007

 Copyright (C) 2007 Free Software Foundation, Inc. <https://fsf.org/>
 Everyone is permitted to copy and distribute verbatim copies
 of this license document, but changing it is not allowed.