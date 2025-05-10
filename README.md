# ✉️ Outlook Automate

## 🛠️ Recursos Necessário
* Pacote Office 365 com o software Classico do Outlook
* [Astral UV](https://docs.astral.sh/uv/)

## 🔰 Como Instalar

```sh
uv pip install -r pyproject.toml
```

## ⌨️ Como Usar
Para utilizar é necessário que você esteja logado no Outlook Classico com a conta pela qual deseja disparar os emails. Observe que você deverá informa-la também nos parametros da CLI, o comando para executar o programa é
```sh
uv run main.py <excel_file> <sender_email>
```
exemplo de uso:
```sh
uv run main.py ./template.xlsx "jonhdoe@example.com"
```

## 📊 Estrutura do arquivo XLSX
É necessário criar uma arquivo XLSX com a seguinte estrutura para realizar os disparos dos emails:
* Uma única worksheet;
* Primeira coluna sendo os emails (se for mais que um deverá ser separado por ";");
* Segunda coluna sendo o assunto;
* Terceira coluna sendo o corpo do email;
* Quarta coluna sendo a pasta que contém os anexos que deseja adicionar na conversa, em caso de não anexo, basta deixa-lo vazio;