# Publicacao de teste

Esta versao esta pronta para teste com poucas pessoas simultaneas.

## Senha do botao Atualizar Dados

Senha local padrao:

```text
play123
```

Na nuvem, configure a variavel de ambiente:

```text
SIMULADOR_UPDATE_PASSWORD
```

## Render

1. Crie uma conta em https://render.com.
2. Suba esta pasta para um repositorio GitHub.
3. No Render, escolha `New` > `Blueprint`.
4. Selecione o repositorio.
5. Configure a variavel `SIMULADOR_UPDATE_PASSWORD`.
6. Publique.

O arquivo `render.yaml` ja define:

- comando de build: `pip install -r requirements.txt`
- comando de start: `python app.py --host 0.0.0.0`
- plano `free` para teste sem disco persistente

No plano gratis, alteracoes no SQLite feitas pelo botao `Atualizar Dados` podem se perder quando o servico dormir, reiniciar ou redeployar. Para teste isso e aceitavel; para producao, use Postgres ou um disco persistente pago.

## Railway

1. Crie um projeto em https://railway.com.
2. Conecte o repositorio GitHub.
3. Configure a variavel `SIMULADOR_UPDATE_PASSWORD`.
4. Use o comando de start:

```text
python app.py --host 0.0.0.0
```

## Local

```powershell
python app.py --host 127.0.0.1 --port 8765
```

Depois acesse:

```text
http://127.0.0.1:8765
```
