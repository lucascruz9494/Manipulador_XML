import time 
import os
import datetime
from dateutil.parser import parse
from xml.dom import minidom
import xml.dom.minidom
import xml.etree.ElementTree as ET
armazena_xml =[] #lista para armazenar o caminho dos arquivos XML

nome_xml =[] #lista para armazenar o nome dos arquivos XML

yes = ["sim","y","yes","Sim","SIM","S","Y","Yes","YES","s"] #lista para input do usuário na escolha da opçao "sim"

no = ["no","No","NO","n","não","Não","NÃO","N","nao","Nao","NAO"] #lista para input do usuário na escolha da opçao "não"

caminhoXml = input("Insira o caminho dos XML :     \n \n \n \n ") #input para o usuário escrever o caminho onde os XML estão , deve ser digitado o caminho completo da pasta , por exemplo "C:\Users\usuario\Desktop\pasta_dos_XML"

path_txt= str(caminhoXml)+"\Cadastro_Produtos.txt" #variável que armazenará o caminho onde ficará o arquivo txt com os valores das tags solicitadas para todos os XML

def create_txt(): #função inicial que servirá para criar o arquivo txt 
    
    try:
        
        write_txt = open(path_txt, "r") #tentará verificar se o arquivo txt já existe .Caso o arquivo txt já exista, irá ser solicitado se deseja sobreescrever o arquivo, ou continuar editando o arquivo já existente
        
        print("Arquivo ja existe \n \n \n")
        
        criaTxt= input("Deseja criar outro, ou continuar escrevendo abaixo da última edição? Digite y para a primeira opção e n para a segunda \n \n \n") 
        
        if criaTxt in yes: 
            
            write_txt = open(path_txt, "w") 
            
            print("Arquivo txt criado \n \n \n")
            
        elif criaTxt in no:
            
            pass
        
        else:
            
            print("Não entendi, repita")
            
            create_txt()
                
    except:
        
        write_txt = open(path_txt, "w")
        
        print("Arquivo txt criado \n \n \n")
        
create_txt() #chama função inicial de criação do txt

for root, dirs, files in os.walk(caminhoXml, topdown=False): #irá verificar todos os arquivos que estão no formato XML no diretório informado no início do script, e irá armazenar o caminho completo na lista a , assim como o nome dos arquivos na lista b
    
    for filename in files:
        
        if filename.endswith('.xml') and filename.startswith(''):
            
            nome_xml.append(filename)
            
            path=os.path.join(root, filename)
            
            armazena_xml.append(os.path.join(root, filename))

conta = 0 #variáveis de apoio para o loop
segue = 0 
count = 0
var = 0
xml_formatado=[] #lista3 para armazenar os xml indentados em arquivos txt ao fim do script
tags = ["cProd","cEAN","xProd","NCM","CFOP","uCom","qCom","vUnCom","vProd","cEANTrib","uTrib","qTrib","vUnTrib","indTot","vTotTrib",
        "orig","CST","modBC","vBC","pICMS","vICMS","CST","vBC","pPIS","vPIS","CST","vBC","pCOFINS","vCOFINS","pDevol","vIPIDevol"] #lista4 com as tags solicitadas referente aos cadastros de produtos
    

def save_archive(ext): #função para salvar os valores das tags em formato .csv e .txt
    global var
    global tags
    global segue
    global lista_xml
    global count
    global armazena_xml
    with open(armazena_xml[conta]+"_PRODUTOS_"+str(ext),"w") as att_file:
        while var <= len(tags)-1:
            att_file.write(str(tags[var])+";")
            var+=1
        att_file.write("\n")


        while segue<=len(lista_xml)-1:
            while count<len(tags):
                att_file.write(lista_xml[segue])
                segue+=1
                count+=1
            count=0
            att_file.write("\n")
        



while conta <= len(armazena_xml)-1: #será iniciado o loop onde os arquivos XML serão lidos e verificados as tags solicitadas
    tags_repetidas = []
    info_repetida = []
    lista_xml=[]
    
    dom = xml.dom.minidom.parse(armazena_xml[conta])  #variável necessária para analisar o xml armazenado na lista a com o ituito de utilizar a biblioteca para indentar o XML
    
    xmldoc = minidom.parse(armazena_xml[conta])
    
    pretty_xml_as_string = dom.toprettyxml() #variável necessária para indentar o arquivo XML
    
    xml_formatado.append(pretty_xml_as_string) #armazenamento do xml indentado na lista3
    
    tree = ET.parse(armazena_xml[conta]) #variável necessária para analisar o xml armazenado na lista a, porém para verificação posterior das tags e atributos
    
    root = tree.getroot()  
    
    text_info = [] 
    
    tags_info = [] 
    
    for elem in root.findall(".//{http://www.portalfiscal.inf.br/nfe}det//"): #laço for para iniciar a verificação das tags solicitadas, encontrando todas as subtags que estão dentro da tag "det"
        
        tag_ide = elem.tag #variável para armazenar o nome da tag
        
        text_ide = elem.text #variável para armazenar o valor dentro da tag
        if text_ide == None or text_ide=="":
            text_ide = "None"
        else:
            pass
        
        tag_ide = tag_ide[36:] #variável para retirar o {http://www.portalfiscal.inf.br/nfe} que está presente em todas as tags, apenas para formatação
        
        if tag_ide == "dhEmi": #Condição if que irá verificar se a tag que contem a data foi encontrada
            
            text_ide = datetime.datetime.strptime(text_ide, "%Y-%m-%dT%H:%M:%S%z").strftime("%d/%m/%Y")  #caso a data seja encontrada, irá ser formatada para ser armazenada no formato solicitado   
        
        text_info.append(text_ide) #armazenamento do valor da tag na lista5
        
        tags_info.append(tag_ide) #armazenamento do nome da tag na lista6

    while segue<=len(tags_info)-1: #laço while para comparar as tags encontradas no XML com as tags armazenadas na lista tags (lista4)
        
        if tags_info[segue] in tags: #caso a tag encontrada no xml seja alguma das solicitadas, irá armazenar no arquivo txt

            
         
            tags_repetidas.append(tags_info[segue])
            info_repetida.append(text_info[segue])
            
            print(tags_repetidas[count]+" : "+info_repetida[count])
                 
            segue+=1
            count+=1
        else:          
            
            segue+=1
    segue=0
    count=0
    with open(path_txt,"a") as att_file: #usado para abrir o arquivo txt e escrever as informações
        while segue<=len(tags_repetidas)-1:
          
            if tags[count] == tags_repetidas[segue]:
         
                att_file.write(str(info_repetida[segue])+";")
                lista_xml.append(str(info_repetida[segue])+";")
                count+=1
                segue+=1
            else:
               
           
                info_repetida.insert(segue,";")
                att_file.write(";")
                lista_xml.insert(segue,"None;")
                count+=1
                segue+=1
            if count>len(tags)-1:
                
             
                att_file.write("\n")
                
                count=0
                
        while count<len(tags):
            info_repetida.insert(segue,";")
            lista_xml.insert(segue,"None;")
            segue+=1
            att_file.write(";")
            count+=1
             
        
                
                                             

    print(nome_xml[conta])
    with open(path_txt,"a") as att_file:
        att_file.write("\n \n")
    var=0
    segue=0
    count=0
    save_archive(".csv")
    var=0
    segue=0
    count=0
    save_archive(".txt")
        
    segue=0
    var=0
    
            
    print(" \n \n \n \n \n \n ")
    
   


    
    conta+=1
    count=0
    
conta=0

    
while conta<=len(xml_formatado)-1: #laço while para escrever os xml indentados em arquivos txt de mesmo nome
    
    salva_txt = open(caminhoXml+"\\"+nome_xml[conta][:-3]+"_INDENTADO_.rtf", "w")
    
    salva_txt.write(xml_formatado[conta])
    
    conta+=1

print("Verifique na pasta "+caminhoXml+" os arquivos com os XML indentados em TXT e o arquivo Cadastro_Produtos.txt com as informações solicitadas \n \n \n \n")
    
fecha_programa = input("Pressione ENTER para encerrar o script \n \n \n \n")

if fecha_programa.startswith(""): #condição if para fechar o script
    quit()




