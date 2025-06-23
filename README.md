# 🟢 Enviador de Relatórios via WhatsApp

Automação Python para enviar relatórios diários por WhatsApp com interface amigável, leitura de planilhas Excel e envio automático com simulação de teclado.

## ✨ Funcionalidades

- Interface gráfica para selecionar contatos
- Leitura automática de planilha Excel
- Mensagem formatada conforme modelo
- Envio automático pelo WhatsApp Web
- Agendamento programado (a cada X minutos)
- Exportável como `.exe` para uso sem Python instalado

## 🚀 Como usar

### Pré-requisitos

- Ter Python instalado (`https://www.python.org`)
- Navegador logado no WhatsApp Web
- Criar estrutura de pastas:

```
    projeto-automacao-python/
    │
    ├── assets/
    │   └── relatorio_diario.xlsx
    │
    ├── dist/
    │   └── main.exe  ← Executável final
    │
    ├── src/
    │   └── main.py
    │
    └── requirements.txt
```

### Instalar dependências

```bash
pip install -r requirements.txt
```

### Executar
```bash
python src/main.py
```

### Gerar .exe (opcional)
```bash
cd src
pyinstaller --noconfirm --onefile --windowed main.py
```

## 📄 Modelo de mensagem
### A mensagem enviada é idêntica ao seguinte formato:
```bash
RELATÓRIO DIÁRIO DE ITEM

Data:
Hora:
Instalador/auxiliar:
Tipo:
Largura:
Altura:
Localização:
Já instalado?
Se não, o que faltou?
Já regulado?
Se não, o que faltou?
Já calafetado?
Se não, o que faltou?
Já finalizado?
Se não, o que faltou?
Faltou algum material/componente?
Descrever qual componente ou perfil ou encaminhar a foto de referência

Documento enviado por [Nome do operador]
```
## 📍 Observações
* WhatsApp Web precisa estar aberto e visível;
* O sistema aguarda alguns segundos e simula o Enter automaticamente;
* Não clique em outras janelas durante o processo de envio.

🤝 Créditos
Desenvolvido por <a href="https://github.com/robertdouglasaimon">Robert</a> com a consultoria do Copilot 💻✨