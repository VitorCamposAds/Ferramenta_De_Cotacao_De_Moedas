# Ferramenta de Cotação de Moedas

Bem-vindo ao repositório da **Ferramenta de Cotação de Moedas**, uma aplicação desenvolvida em Python utilizando Tkinter para interface gráfica. Este projeto é ideal para desenvolvedores que desejam explorar a integração de APIs, manipulação de dados com Pandas e a construção de interfaces amigáveis.

## 🛠️ Tecnologias Utilizadas

- **Python**: Linguagem de programação principal.
- **Tkinter**: Biblioteca para criação de interfaces gráficas.
- **tkcalendar**: Para seleção de datas.
- **Pandas**: Manipulação de dados em formato de tabela (DataFrames).
- **NumPy**: Operações numéricas.
- **Requests**: Realização de requisições HTTP para APIs.
- **Excel**: Exportação de dados.

## 📊 Funcionalidades

### 1. Obtenção de Cotações
- O usuário pode buscar a cotação de uma moeda específica para uma data escolhida.
- A aplicação verifica a validade da moeda selecionada.

### 2. Atualização em Massa
- É possível carregar um arquivo Excel contendo múltiplas moedas e atualizar suas cotações em um intervalo de datas.
- A aplicação gera um novo arquivo Excel com as cotações atualizadas.

### 3. Interface Amigável
- A interface foi projetada para ser intuitiva, permitindo uma fácil navegação entre as funcionalidades.
- Utiliza caixas de diálogo para seleção de arquivos e mensagens informativas.

## 🔗 Como Usar

1. **Clone o Repositório**
   ```bash
   git clone https://github.com/seu_usuario/Ferramenta-Cotacao-Moedas.git
   cd Ferramenta-Cotacao-Moedas
   ```

2. **Instale as Dependências**
   Certifique-se de ter o Python instalado e, em seguida, execute:
   ```bash
   pip install -r requirements.txt
   ```

3. **Execute a Aplicação**
   ```bash
   python seu_script.py
   ```

4. **Interaja com a Interface**
   - Selecione uma moeda e uma data para pegar a cotação específica.
   - Para atualizar várias moedas, carregue um arquivo Excel e defina o intervalo de datas.

## 💡 Contribuição

Se você tem ideias para melhorias ou novas funcionalidades, sinta-se à vontade para abrir uma *issue* ou enviar um *pull request*. A colaboração é sempre bem-vinda!


## 🌟 Agradecimentos

- Agradecer-se-á à API AwesomeAPI por fornecer as cotações de moedas de forma acessível e fácil de usar.

---

Sinta-se à vontade para explorar e aprimorar este projeto!
