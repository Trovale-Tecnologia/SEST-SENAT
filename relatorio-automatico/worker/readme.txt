Para executar o script Python que coordena o preenchimento de relatório mensal ao SEST SENAT é necessário seguir os seguintes passos com atenção:

ATENÇÃO! 
ESSA PARTE SÓ PRECISA SER EXECUTADA UMA ÚNICA VEZ PARA FINS DE CONFIGURAÇÃO DO AMBIENTE DE EXECUÇÃO DE SCRIPTS 

========================================================================================================================

    1. Instalar Python 3.11 ou mais atualizado (https://www.python.org/downloads/);
        -> Para verificar se o python foi instalado corretamente execute o seguinte comando: python --version

    2. Abrir o CMD, PowerShell ou qualquer interface de linha de comando e navegar até o diretório "/worker" que contém o arquivo '"main.py"'

    3. Instalar o gerenciador de dependencias do projeto:
        -> executar no terminal os seguintes comandos um após o outro:
            -: curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py (esse comando obtém o  PIP, que é gerenciador de pacotes do python e o utilizaremos para instalar as bibliotecas)
            -: python get-pip.py (executa o instalador do PIP para gerenciamento de pacotes)
            -: pip --version (para verificar se o gerenciador de pacotes foi instalado corretamente)
    
    4. Instalar as dependencias do projeto:
        -> executar no terminal os seguinte comando um após o outro:
            -: pip install pandas
            -: pip install openpyxl
    
    5. Executar o arquivo de execução "main.py" dentro do diretório "/worker  com o comando:
        -:  python main.py 
        
            P.S.: (dependendo da sua versão de python pode ser necessário utilizar o comando python3, python.exe ou python3.exe para executar o arquivo de script)


========================================================================================================================
