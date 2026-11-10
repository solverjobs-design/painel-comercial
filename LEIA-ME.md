# Painel comercial automático (Dinâmica)

Todo dia às 6h (Brasília) o GitHub lê a planilha do SharePoint, gera o painel e publica num link fixo.
Ninguém precisa deixar computador ligado. O time só continua atualizando a planilha.

## Antes de começar: 1 arquivo fixo no SharePoint
O painel lê **um arquivo só**. O time deve editar sempre o mesmo arquivo (mesmo nome), sem criar
"Status_Comercial_9", "_10". Se o arquivo for substituído por outro nome, o link para de funcionar.
Os títulos dos blocos na coluna B (Clientes no radar, Solicitação de orçamento, Orçamento Enviado,
Visitação, Contrato /Aditivo, Contrato Assinado, Check in já realizados, Clientes perdidos) não podem mudar.

## Passo a passo (uma vez só, feito com as contas da Dinâmica)

**1. Link da planilha**
No SharePoint, abra o arquivo, clique em Compartilhar, escolha "Qualquer pessoa com o link", permissão
"Pode visualizar", e copie o link. Se essa opção não aparecer, o TI da Dinâmica bloqueou links abertos;
avise para trocarmos para o método com aplicativo do Microsoft Entra.

**2. Repositório no GitHub**
Crie um repositório chamado `painel-comercial` e marque **Public** (o GitHub Pages gratuito exige).
Envie todo o conteúdo desta pasta, inclusive a pasta oculta `.github`. O repositório guarda só código
e a contagem por etapa (historico.json); nunca nomes de clientes nem a planilha.

**3. Guardar o link como segredo**
No repositório: Settings > Secrets and variables > Actions > New repository secret.
Nome: `PLANILHA_URL`. Valor: o link do passo 1. Não coloque o link no config.json, porque o repositório é público.

**4. Ligar o GitHub Pages**
Settings > Pages > Source: **GitHub Actions**.

**5. Primeira atualização**
Aba Actions > "Atualizar painel" > Run workflow. Em 1 a 2 minutos o painel fica em
`https://NOME-DA-CONTA.github.io/painel-comercial/`. Esse é o link fixo.

## Dia a dia
- Atualiza sozinho todo dia às 6h (o GitHub pode atrasar alguns minutos).
- Para atualizar na hora: Actions > Atualizar painel > Run workflow.
- Para parar: Actions > Atualizar painel > menu "..." > Disable workflow.
- Se a planilha estiver quebrada ou o link parar de funcionar, a atualização falha, o GitHub manda e-mail
  ao dono do repositório e **o painel antigo continua no ar**. O motivo aparece no log da execução.
- O GitHub pode pausar agendamentos de repositório público sem atividade por 60 dias. Se o painel parar
  de atualizar, clique em Enable workflow na aba Actions.

## Atenção: o painel é público
Quem tiver o link do painel vê os nomes dos clientes e as quantidades de pessoas. Não divulgue o link fora
da empresa. Se isso incomodar, a alternativa é publicar na Cloudflare com senha.

## Arquivos
- `config.json`: aba lida da planilha (Planilha1). O link vem do segredo PLANILHA_URL.
- `gerar.py`: lê a planilha, confere os dados e gera `site/index.html`. Teste local: `python gerar.py arquivo.xlsx`.
- `modelo.html`: visual do painel.
- `.github/workflows/atualizar.yml`: agendamento diário às 6h e publicação.
- `historico.json`: contagem por etapa a cada dia (alimenta o gráfico de evolução).
