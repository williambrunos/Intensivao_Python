# Dia 1

O projeto do dia 1 consiste em uma automação de um processo corriqueiro em uma empresa: baixar uma base de dados, calcular os indicadores e enviar um resumo para uma lista de e-mails.

## Dependências:
- pyautogui
- pyperclip
- pandas
- time

## Funcionamento

Pyautogui é uma lib de automação de teclado e mouse e é com ela que majoritariamente realizamos a automação. Pyperclip funciona para escrever caracteres especiais com pyautogui. Pandas funciona para realizar o cálculo dos indicadores da base de dados e time para dar uma pausa entre cada execução do código.

O script em python faz:
- Abre o google chrome pelo windows.
- Escreve o link do drive na barra de busca e navega até ele.
- Entra na pasta na qual se encontra a base de dados e faz o seu download.
- Copia a base dados no diretório do script.
- Realiza o cálculo dos indicadores.
- Envia um resumo para uma lista de e-mails.

## Observação

A automação é feita com base no SO Windows e com um monitor específico. Para realizar a mesma automação na sua máquina, você deve obter as posições de cada elemento na qual deseja navegar usando a função:

``pyautogui.position()``

Essa função retorna uma tupla com as posições x e y do seu mouse naquele instante em que fora chamada.
Você pode utilizar a função ``sleep`` da biblioteca time para ajudar nessa tarefa.
