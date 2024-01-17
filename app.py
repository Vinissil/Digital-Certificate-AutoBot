'''

Quero ver a possibilidade de criar um programa usando o Python para automatizar enviando os dados da planilha para preencher os campos mutáveis no certificado padrão.

Tipo nome do curso, nome participante, tipo de participação, data do inicio, data do final, carga horária, data da emissão do certificado e as assinaturas do Gestor Geral, do Coordenador e do aluno.

Pegar dados(texto) -> sobrepor de uma imagem(certificado padrão)

#Pegar os dados da planilha
Tais como nome do curso, nome participante, tipo de participação, data do inicio, data do final, carga horária, data da emissão do certificado

#tranferir para a imagem do certificado

I want to see the possibility of creating a program using Python to automate sending spreadsheet data to populate the changeable fields in the default certificate.

Like course name, participant name, type of participation, start date, end date, workload, certificate issuance date and the signatures of the General Manager, the Coordinator and the student.

Get data (text) -> overlay from an image (standard certificate)

#Get the data from the spreadsheet
Type course name, participant name, type of participation, start date, end date, workload, certificate issuance date

#transfer to certificate image

'''

#Pegar os dados da planilha #Get the data from the spreadsheet

#Get the data from the spreadsheet
#importar bibliotecas usadas para rodar o codigo openpyxl(import openpyxl) para abrir o arquivo excel e o pillow (from PIL import Image, ImageDraw, ImageFont) para abrir a imagem do certificado #import libraries used to run the code openpyxl(import openpyxl) to open the excel file and the pillow(from PIL import Image, ImageDraw, ImageFont) to open the certificate image
import openpyxl
from PIL import Image, ImageDraw, ImageFont

#criar as variaveis para abrir a planilha e a aba da planilha #create the variables to open the spreadsheet and the spreadsheet tab
workbookStudents = openpyxl.load_workbook('planilha_alunos.xlsx')
sheetStudents = workbookStudents['Sheet1']

#colocar um indice para cada linha da planilha e colocar um liite minimo e maximo para gerar os certificados durante o periodo de teste #put an index for each row of the spreadsheet and put a minimum and maximum limit to generate the certificates during the test period
for index, line in enumerate(sheetStudents.iter_rows(min_row = 2)):

    #a célula que contém a info que precisamos #each cell that contains the information we need
    courseName = line[0].value #nome do curso #course name
    participantName = line[1].value #nome do participante #participant name
    typeParticipation = line[2].value #tipo de participação #type of participation
    startDate = line[3].value #data do inicio #start date
    endDate = line[4].value #data do final #end date
    workload = line[5].value #carga horária #workload
    certificateIssuanceDate = line[6].value #data da emissão do certificado #certificate issuance date

    #tranferir os dados da planilha para a imagem do certificado #transfer the data from the spreadsheet to the certificate image
    #definindo fonte a ser usada #defining font to be used
    fontName = ImageFont.truetype('./tahomabd.ttf', 90)
    generalFont = ImageFont.truetype('./tahoma.ttf', 80)
    dataFont = ImageFont.truetype('./tahoma.ttf', 55)

    #abrir a imagem do certificado #open the certificate image
    image = Image.open('./certificado_padrao.jpg')
    draw = ImageDraw.Draw(image)

    #escrever na imagem do certificado #write on the certificate image
    draw.text((1020, 827), participantName, fill = 'black', font=fontName)
    draw.text((1060, 950), courseName, fill = 'black', font = generalFont)
    draw.text((1435, 1065), typeParticipation, fill = 'black', font=generalFont)

    #Para o programa ler datas e carga horária, é preciso mudar a class de número para string #For the program to read dates and workload, you need to change the number class to string
    draw.text((1480, 1190), str(workload), fill = 'black', font = generalFont)

    #assim é feito com a data de inicio e de termino #this is done with the start and end date
    draw.text((750, 1770), str(startDate), fill = 'blue', font = dataFont)
    draw.text((750, 1930), str(endDate), fill = 'blue', font = dataFont)
    
    draw.text((2220, 1930), certificateIssuanceDate, fill = 'blue', font = dataFont)
    

    #salvar a imagem do certificado #save the certificate image
    #rodar o codigo para gerar o primeiro certificado #run the code to generate the first certificate
    image.save(f'./certificate_img/ {index} {participantName} certificate.png')
    
