# Portal_Relatorio_Estagios_
Sistema de apoio a estagiários: Gerador Automático de Relatórios de Estágio com Python , Streamlit e Google Forms

Objetivo

Este pequeno sistema foi criado para faciliatar e automatizar o processo de escrita do relatório de estágio.
É uma soluçãopensada para que os estagiários não se esqueçam de nenhum detalhe/ tarefa que foi realizada durante todo o tempo de estágio, ficando tudo registado : Tarefas, Aprendizagens e Dificuldades ultrapassadas.
Assim o foco é na experiência eliminando a preocupação constantes da docuemntação.


Fluxo de Dados

1. Recolha de dados : Registos de atividades submetidas semanalmente através do Google Forms
2. Base de Dados : Respostas aramazenadas automaticamnete no Google Sheets
3. Proceseso d & Exportação : A aplicação Streamlit lê a folha de cálculo, aplica filtros pelo o utilizador (Pin segurança) e gera o documento Word formatado.


Tecnologias Utilizadas

- Linguagem: Python
- Frontend: Streamlit
- Manipulação de Dados: Pandas
- Geração de Documentos: Python-docx
- Backend/Data: Google Sheets Api
- Segurança: Autenticação simples (PIN)


Instalação e Excução Local

Para correr o projeto localmente: 
1- clona o repositório: Git clone 
2- Cria uma ambiente virtual e instala as dependências: 
  python -m ven 
  pip install -r requirements.txt
3- Executar a App: streamlit [nome do ficheiro].py 



<img width="1271" height="487" alt="image" src="https://github.com/user-attachments/assets/30f6abb2-5aa3-4673-b8a4-d5cfe9e82c6b" />

