#!/usr/bin/env python3
"""Ordena a aba BM - GCP Atividades - DP pela lógica da arquitetura.

O cronograma escrito na planilha é preservado e não define a ordem.
Lacunas entram só quando o documento ou a dependência de um card as nomeia.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = "/workspace/BM-GCP-Atividades-DP-ordem-temporal.xlsx"

# origem, etapa, card_id, descricao, complexidade, dependencia, cronograma,
# responsavel, observacoes, base, posicao, url, linha_original
# linha_original 0 = lacuna, não veio da planilha

CARD = "Card da planilha"
GAP_DOC = "Lacuna — documentação"
GAP_DEP = "Lacuna — dependência citada no card"
MARCO = "Marco pedido, sem card"

ROWS = [
    (
        CARD,
        "1. HML — base já concluída",
        "424950",
        "[DP] Provisionar ambiente GKE HML",
        "Concluído",
        "",
        "",
        "",
        "",
        "Status Concluído. Único card de ambiente, nome HML. A documentação cita o cluster Azure de homologação azr-hml02, namespaces busca, dataprovider-data e portal-cadastro.",
        "Ponto de partida. O pedido é sair da homologação.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424950/details",
        40,
    ),
    (
        CARD,
        "1. HML — base já concluída",
        "451661",
        "[DP] Migração GCP | Criar banco de dados",
        "Concluído",
        "",
        "",
        "",
        "",
        "Status Concluído. O card de migrar o SQL Server está abortado. A documentação descreve o Corporate (HML, master, réplica e staging) como o banco que serve as APIs.",
        "O banco existe. A carga desse banco não tem card e entra mais adiante, antes do teste funcional.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451661/details",
        39,
    ),
    (
        CARD,
        "1. HML — base já concluída",
        "451684",
        "[DP] Migração GCP | Redis",
        "Concluído",
        "Os pods precisam estar de pé para testar",
        "",
        "Back End",
        "",
        "Status Concluído. No diagrama, busca-cache fica entre o Corporate e as APIs. A documentação lista dois Redis de 6 GB: DP Cache PRD e Busca. A planilha não separa HML e PRD.",
        "Concluído. O teste que a dependência pede espera os pods das APIs.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451684/details",
        35,
    ),
    (
        CARD,
        "1. HML — base já concluída",
        "476416",
        "Continuação - Criação de Workloads do projeto BUSCA -> datp-searcher-api-commons",
        "Concluído",
        "",
        "",
        "",
        "",
        "Status Concluído. O card 470112, da criação do mesmo workload, segue com observação 'atuando'.",
        "Concluído. O 470112 permanece na fila das APIs.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/476416/details",
        30,
    ),
    (
        CARD,
        "2. HML — antes de subir workload",
        "424949",
        "[DP] Definir conectividade",
        "",
        "",
        "",
        "",
        "",
        "Único card de conectividade. Os workloads de busca registram que a conectividade por IP está errada. A seção de rede do documento (VNet, subnet, VPN, DNS, firewall) está vazia. O diagrama mostra as fontes K2/Gliese, base federada, MDM, GSC, hospitais e MV.",
        "Antes dos workloads. Sem isso a aplicação sobe apontando para IP, que a própria planilha registra como errado.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424949/details",
        43,
    ),
    (
        CARD,
        "2. HML — antes de subir workload",
        "424942",
        "[DP] Adequação de pipelines CI/CD",
        "Alta",
        "Infra, JulianoTelli",
        "",
        "Infra",
        "telli",
        "Card de esteira, sem data. Nenhum workload cita este número como dependência.",
        "Em paralelo com a conectividade. A planilha não diz que este card bloqueia a criação do workload.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424942/details",
        31,
    ),
    (
        CARD,
        "5. HML — migrar o Airflow da Azure para a GCP",
        "478326",
        "Spike -> Airflow na versao mais nova do airflow",
        "Alta",
        "Victor, subir um docker e testar ambienteLocal",
        "20/10/26",
        "Back End",
        "Aqui já vai ser possivel pegar imcompatibilidade de versao",
        "O Airflow de partida já está na Azure. A dependência deste card é docker local. A documentação registra a versão 2.2.5 como baixa demais para a versão do Google.",
        "Primeira linha desta frente, em paralelo com as APIs da etapa 3. A spike existe por causa da versão, antes de migrar o Airflow da Azure para a GCP. Carga, Blob e portal esperam essa migração. A data 20/10/26 não posiciona o resto.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/478326/details",
        2,
    ),
    (
        CARD,
        "5. HML — migrar o Airflow da Azure para a GCP",
        "",
        "Criar historia adequação airflow, vinculada ao fim da poc do Victor => necessario devido versao de airflow",
        "Alta",
        "as adequações levarão tempo, pode ser a maior data.",
        "entre março e abril de 2027",
        "",
        "",
        "Linha da planilha, sem Card ID. É a migração do Airflow que já roda na Azure para a GCP. A adequação existe por causa da versão. Portal, Blob e backoffice citam Airflow como dependência. Os workloads com prd no nome ficam na produção.",
        "Depois da spike. A janela de março–abril de 2027 fica registrada e não empurra esta linha para o fim. Blob e portal ficam depois porque dependem do Airflow já na GCP.",
        "",
        13,
    ),
    (
        CARD,
        "3. HML — APIs e gateway, em paralelo com o Airflow",
        "470110",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-brand",
        "Baixa",
        "Backend Hj a conectividade está por IP, isso é errado",
        "20/10/26",
        "Back End",
        "",
        "Uma das seis APIs do documento: exam, healthinsurance, commons, agendamento, local e brand. Este card não cita Airflow nem Cognitive Search.",
        "Depois da conectividade. Não espera o Airflow. A data 20/10/26 foi ignorada na ordenação.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470110/details",
        3,
    ),
    (
        CARD,
        "3. HML — APIs e gateway, em paralelo com o Airflow",
        "470109",
        "Criação de Workloads do projeto BUSCA -> datp-infra-gateway -> (ZUUL)",
        "Média",
        "Backend Hj a conectividade está por IP, isso é errado",
        "27/10/26",
        "Back End",
        "",
        "No diagrama o gateway aparece como dp-zuul-proxy, ao lado de autorização, rate limit e WAF. Não há card separado para WAF, rate limit ou API key.",
        "Junto das APIs, depois da conectividade. A documentação não separa o gateway da entrada das APIs.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470109/details",
        4,
    ),
    (
        CARD,
        "3. HML — APIs e gateway, em paralelo com o Airflow",
        "470118",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-local",
        "Média",
        "Backend Hj a conectividade está por IP, isso é errado",
        "07/11/26",
        "Back End",
        "",
        "API local do documento. Não cita Airflow nem Cognitive Search.",
        "Depois da conectividade. Data da planilha ignorada na ordenação.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470118/details",
        5,
    ),
    (
        CARD,
        "3. HML — APIs e gateway, em paralelo com o Airflow",
        "470115",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-healthinsurance",
        "Alta",
        "Backend Hj a conectividade está por IP, isso é errado",
        "20/11/26",
        "Back End",
        "",
        "API de convênio do documento. Não cita Airflow nem Cognitive Search.",
        "Depois da conectividade. Data da planilha ignorada na ordenação.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470115/details",
        6,
    ),
    (
        CARD,
        "3. HML — APIs e gateway, em paralelo com o Airflow",
        "470113",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-elegibility",
        "Alta",
        "Backend Hj a conectividade está por IP, isso é errado",
        "29/11/26",
        "Back End",
        "",
        "O documento e o diagrama nomeiam busca-agendamento. A planilha nomeia elegibility. Não há outro card com o nome agendamento.",
        "Permanece nesta fila porque o card não cita Airflow nem Cognitive Search. Não foi tratado como a API de agendamento.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470113/details",
        7,
    ),
    (
        CARD,
        "3. HML — APIs e gateway, em paralelo com o Airflow",
        "470112",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-commons",
        "Baixa",
        "Backend Hj a conectividade está por IP, isso é errado",
        "",
        "Back End",
        "atuando",
        "API commons do documento. A continuação 476416 está concluída. Este card segue 'atuando'.",
        "Na mesma fila das outras APIs de busca que não citam índice.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470112/details",
        17,
    ),
    (
        CARD,
        "3. HML — APIs e gateway, em paralelo com o Airflow",
        "424944",
        "[DP] Migração de aplicações, APIs",
        "Baixa",
        "testar conectividade",
        "15/12/26",
        "Back End",
        "",
        "A dependência é testar conectividade, não teste funcional.",
        "Logo depois dos workloads desta fila. Confirma que a aplicação enxerga o destino. O teste funcional espera dado, índice e portal.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424944/details",
        8,
    ),
    (
        CARD,
        "3. HML — front e back do DP, em paralelo com o Airflow",
        "469049",
        "[DP] Migração GCP | Migrar Backend -> front do DP (adaptações)",
        "Alta",
        "Ajuste de codigo, versao, eventual erro",
        "",
        "Back End",
        "",
        "A documentação aponta o repositório datp-backoffice-backend. O card não cita Airflow. O diagrama tem dp-backoffice-api-prd.",
        "Sobe com as APIs. Não espera carga nem índice. O teste do frontend espera a integração.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/469049/details",
        15,
    ),
    (
        CARD,
        "3. HML — front e back do DP, em paralelo com o Airflow",
        "469051",
        "[DP] Migração GCP | Migrar Frontend  portal do DP (adaptações)",
        "Média",
        "Ajuste de codigo, versao, eventual erro",
        "",
        "Back End",
        "só é possivel testar o front com tudo integrado.",
        "A documentação cita o frontend em dp-corporate-frontend e o repositório datp-backoffice-frontend. O card 470121, dp-corporate-frontend, está cancelado. A observação deste card fala do teste, que é o 451689.",
        "A adaptação entra aqui. O teste do front fica no fim da HML, com o 451689.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/469051/details",
        14,
    ),
    (
        CARD,
        "4. HML — índice e exam, em paralelo com o Airflow",
        "451682",
        "[DP] Migração GCP | Index (cognitive search > vertex - adaptar na aplicação)",
        "Alta",
        "Maior Risco, descoberta",
        "",
        "Felipe Paixão | Back End",
        "",
        "No diagrama, busca-search fica entre o Corporate e as APIs. No documento, Cognitive Search é dependência externa, criticidade alta, e 'pode desacoplar' está TBD. O card aponta Vertex e está em descoberta. A planilha só declara a dependência na API de exam e no dp-backend-api cancelado.",
        "Antes da API de exam. As outras APIs não foram colocadas atrás deste card: a planilha não diz que elas leem o índice.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451682/details",
        41,
    ),
    (
        CARD,
        "4. HML — índice e exam, em paralelo com o Airflow",
        "470114",
        "Criação de Workloads do projeto BUSCA -> datp-searcher-api-exam",
        "Alta",
        "Backend Hj a conectividade está por IP, isso é errado",
        "",
        "Back End",
        "depende cognitiveSearch",
        "API de exam do documento. Observação: depende de Cognitive Search. Também registra conectividade por IP errada.",
        "Depois do 451682 e depois da conectividade.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470114/details",
        18,
    ),
    (
        GAP_DOC,
        "6. HML — dado, depois do Airflow e do backend",
        "",
        "Carga, no banco de HML, dos dados de exame, unidade e convênio que hoje estão no Corporate",
        "",
        "",
        "",
        "",
        "",
        "O DP serve exame, unidade e convênio para NAV e CDC. O diagrama leva as fontes ao Corporate, e o Corporate às APIs. Existe CORPORATE_HOM na documentação. O card 451661 cria o banco e está concluído. O card 451676, migrar SQL Server, está abortado. Não há card de carga.",
        "Antes do teste funcional e depois da migração do Airflow para a GCP. Se o banco criado já tiver esses dados, esta linha sai.",
        "",
        0,
    ),
    (
        CARD,
        "6. HML — dado, depois do Airflow e do backend",
        "451688",
        "[DP] Migração GCP | Blob Storage (adaptações na aplicação)",
        "Média",
        "Criado, depende de airFlow, backend e frontEnd",
        "03/02/27",
        "Back End | Dados",
        "",
        "A documentação registra 269 GB de Blob. A dependência do card pede Airflow, backend e frontend. A data 03/02/27 não define a posição.",
        "Depois do Airflow de HML e depois do backend e do frontend do DP. A planilha não diz se o download D+1 lê este Blob; por isso o D+1 não foi colocado obrigatoriamente depois.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451688/details",
        10,
    ),
    (
        GAP_DEP,
        "6. HML — dado, depois do Airflow e do backend",
        "",
        "Cargas de cadastro",
        "",
        "",
        "",
        "",
        "",
        "Dependência escrita nos cards 470165 e 470100: 'Depende do airFlow pronto, cargas cadastro e D+1 (dowloads Cadastro)'. O diagrama tem o Portal de Cadastros como consumidor e, no desenho GCP, Excel do Portal Cadastro como fonte.",
        "Antes dos workloads de portal que citam essa carga. Não há card, responsável nem prazo.",
        "",
        0,
    ),
    (
        GAP_DEP,
        "6. HML — dado, depois do Airflow e do backend",
        "",
        "Downloads D+1 de cadastro",
        "",
        "",
        "",
        "",
        "",
        "Mesma dependência dos cards 470165 e 470100. Não há card só para o D+1.",
        "No mesmo bloco das cargas de cadastro. A ordem entre D+1 e Blob não está escrita.",
        "",
        0,
    ),
    (
        CARD,
        "7. HML — portal, depois da carga que ele cita",
        "468966",
        "Criação de Workloads do projeto BUSCA -> datp-backoffice-backend",
        "Alta",
        "Ajuste de codigo, versao, eventual erro",
        "",
        "Back End",
        "depende do airFlow",
        "Observação do card: depende do Airflow. Não cita a carga de cadastro.",
        "Depois do Airflow de HML. Antes do teste integrado.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/468966/details",
        22,
    ),
    (
        CARD,
        "7. HML — portal, depois da carga que ele cita",
        "470165",
        "Criação de Workloads do projeto PORTAL-CADASTRO -> portal-cadastro-web",
        "Média",
        "Depende do airFlow pronto, cargas cadastro e D+1 (dowloads Cadastro)",
        "",
        "Back End | Dados | Infra",
        "depende do airFlow",
        "A documentação cita o frontend do portal de cadastro e o repositório datp-backoffice-frontend-portal. O card desse repositório, 468985, está cancelado.",
        "Depois de Airflow, cargas de cadastro e D+1, como a dependência escreve.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470165/details",
        19,
    ),
    (
        CARD,
        "7. HML — portal, depois da carga que ele cita",
        "470100",
        "Criação de Workloads do projeto BUSCA -> datp-backoffice-frontend",
        "Média",
        "Depende do airFlow pronto, cargas cadastro e D+1 (dowloads Cadastro)",
        "",
        "Back End | Dados | Infra",
        "depende do airFlow",
        "Mesma dependência do 470165.",
        "Junto do portal-cadastro-web.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470100/details",
        21,
    ),
    (
        CARD,
        "7. HML — portal, depois da carga que ele cita",
        "451681",
        "[DP] Migração GCP | Migrar Backend -> portal de cadastro (adaptações)",
        "Alta",
        "Ajuste de codigo, versao, eventual erro",
        "10/02/27",
        "Back End",
        "",
        "Adaptação do backend do portal. A data 10/02/27 não define a posição. O card não cita Airflow no texto; o workload do portal cita.",
        "No bloco do portal, junto dos workloads que a planilha amarra à carga. A data de fevereiro ficou para trás da adequação do Airflow de propósito.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451681/details",
        12,
    ),
    (
        CARD,
        "7. HML — portal, depois da carga que ele cita",
        "451677",
        "[DP] Migração GCP | Migrar Frontend  Protal de cadastro (adaptações)",
        "Alta",
        "Ajuste de codigo, versao, eventual erro",
        "07/02/27",
        "Back End",
        "só é possivel testar o front com tudo integrado.",
        "A observação condiciona o teste à integração. A data 07/02/27 não define a posição.",
        "A adaptação fica neste bloco. O teste do front é o card 451689.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451677/details",
        11,
    ),
    (
        CARD,
        "8. HML — teste funcional, e em seguida carga",
        "451689",
        "[DP] Migração GCP | Testes Integrados, Homologação & Plano de Rollback",
        "Sem Definição de Escopo",
        "??",
        "",
        "",
        "",
        "Único card de teste. Não cita teste de carga. Sem definição de escopo. O template do documento lista teste funcional, performance, segurança e DR, todos sem conteúdo. Os fronts 451677 e 469051 dizem que só testam com tudo integrado.",
        "Fecha o que a planilha chama de homologação. Produção só começa depois deste card e do teste de carga.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451689/details",
        36,
    ),
    (
        MARCO,
        "8. HML — teste funcional, e em seguida carga",
        "",
        "Teste de carga em HML",
        "",
        "",
        "",
        "",
        "",
        "Pedido desta ordenação: ter tudo em HML para testes funcionais e de carga antes de planejar produção. O documento cita Performance no template de testes e não define volume, API nem meta. Nenhum card descreve o teste.",
        "Depois do teste funcional. Sem cenário, volume ou API. Não é um card do BusinessMap.",
        "",
        0,
    ),
    (
        GAP_DOC,
        "10. Produção — planejar depois do teste de HML",
        "",
        "GKE de produção",
        "",
        "",
        "",
        "",
        "",
        "A documentação cita o cluster aks-prd-cadastro01, namespaces busca, dataprovider-data e portal-cadastro. O único card de GKE é o de HML, concluído.",
        "Primeiro item de produção porque a HML já tem cluster e a produção, no documento, também tem. Não há card.",
        "",
        0,
    ),
    (
        CARD,
        "10. Produção — planejar depois do teste de HML",
        "470150",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-worker-statefulset",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "Nome com prd. Data escrita 09/10/26. A dependência é infra, não a POC do Airflow.",
        "Produção, pelo nome. A data 09/10/26 não trouxe este card para antes da HML.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470150/details",
        23,
    ),
    (
        CARD,
        "10. Produção — planejar depois do teste de HML",
        "470149",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "Nome com prd. Data escrita 09/10/26.",
        "No bloco de produção, na ordem em que a planilha lista os workloads de Airflow.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470149/details",
        24,
    ),
    (
        CARD,
        "10. Produção — planejar depois do teste de HML",
        "470148",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-sync-users",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "Nome com prd. Data escrita 09/10/26.",
        "No bloco de produção, na ordem original.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470148/details",
        25,
    ),
    (
        CARD,
        "10. Produção — planejar depois do teste de HML",
        "470146",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-scheduler",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "Nome com prd. Data escrita 09/10/26.",
        "No bloco de produção, na ordem original.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470146/details",
        26,
    ),
    (
        CARD,
        "10. Produção — planejar depois do teste de HML",
        "470145",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-flower",
        "Média",
        "infra managedAirFlow - infra => sem data depende de infra",
        "09/10/26",
        "Infra",
        "infra",
        "Nome com prd. Data escrita 09/10/26.",
        "No bloco de produção, na ordem original.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470145/details",
        27,
    ),
    (
        CARD,
        "10. Produção — planejar depois do teste de HML",
        "468978",
        "Criação de Workloads do projeto DATAPROVIDER-DATA -> dp-data-stack-airflow-prd-db-migrations",
        "Média",
        "infra managedAirFlow - infra",
        "09/10/26",
        "Infra",
        "infra",
        "Nome com prd. Data escrita 09/10/26.",
        "No bloco de produção, na ordem original.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/468978/details",
        28,
    ),
    (
        CARD,
        "10. Produção — planejar depois do teste de HML",
        "451683",
        "[DP] Migração GCP | Event Hub para Pub/Sub",
        "Média",
        "Ultima demanda, por ultimo -> falar depois",
        "",
        "BD",
        "pode ficar por ultimo",
        "A documentação lista o Event Hub data-provider-prd. O desenho GCP tem Pub/Sub na ingestão. A planilha manda deixar por último. Nenhuma API cita Event Hub como dependência.",
        "No fim da produção, antes do cutover. Fora do teste funcional de HML, como o card pede.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451683/details",
        29,
    ),
    (
        CARD,
        "10. Produção — planejar depois do teste de HML",
        "451691",
        "[DP] Migração GCP | Cutover & Go-Live",
        "Sem Definição de Escopo",
        "??",
        "",
        "",
        "",
        "Sem definição de escopo. O template do documento deixa em branco big bang ou gradual, DNS e janela.",
        "Última linha de produção. Sem escopo preenchido.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451691/details",
        34,
    ),
    (
        CARD,
        "10. Documentado, fora deste caminho",
        "478328",
        "BigQuery => substituir dataBricks",
        "Alta",
        "Galerinha de dados DP",
        "29/01/27",
        "Dados",
        "Levar toda a orquestração do dataBricks para o BigQuery",
        "Card mantido com o texto original. A decisão registrada é levar o Databricks para a GCP. O diagrama AS IS não desenha o Databricks no caminho até as APIs. O documento lista os workspaces de HML e de PRD e não descreve jobs. O card 451674, converter ETLs, está cancelado. O desenho GCP é uma plataforma genérica com BigQuery bronze, silver e gold; a matriz de estratégia do documento está em branco.",
        "Fora da fila de HML e fora da produção. Nenhum outro card o cita. A data 29/01/27 não o coloca entre dezembro e fevereiro.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/478328/details",
        9,
    ),
    (
        GAP_DOC,
        "9. HML — Databricks na GCP, antes da produção",
        "",
        "Workspaces Databricks de HML e de PRD na GCP",
        "",
        "",
        "",
        "",
        "",
        "A documentação lista o workspace de PRD adb-7665073515379548 e o de HML adb-887971986833958. Não há jobs, notebooks nem tabelas de saída. Não há card para levar esses workspaces à GCP. O card 478328, que pede BigQuery, está na aba de discussão.",
        "Entra porque a decisão é levar o Databricks para a GCP e a documentação comprova que os workspaces existem. Fica depois do teste funcional: o diagrama até as APIs não passa pelo Databricks, então o teste não espera esta linha. Produção começa depois dela.",
        "",
        0,
    ),
    (
        CARD,
        "Encerrada pela equipe",
        "470119",
        "Criação de Workloads do projeto BUSCA -> dp-backend-api",
        "Cancelado",
        "",
        "",
        "",
        "depende cognitiveSearch",
        "Complexidade Cancelado. A observação cita Cognitive Search.",
        "Cancelado pela equipe. Os backends 469049 e 468966 continuam no caminho. O documento não descreve este card além da observação de Cognitive Search.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470119/details",
        16,
    ),
    (
        CARD,
        "Encerrada pela equipe",
        "468985",
        "Criação de Workloads do projeto PORTAL-CADASTRO -> datp-backoffice-frontend-portal",
        "Cancelado",
        "",
        "",
        "",
        "depende do airFlow",
        "Complexidade Cancelado.",
        "Movida para a aba Discutir com a equipe.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/468985/details",
        20,
    ),
    (
        CARD,
        "Encerrada pela equipe",
        "470121",
        "Criação de Workloads do projeto BUSCA -> dp-corporate-frontend",
        "Cancelado",
        "",
        "",
        "",
        "",
        "Complexidade Cancelado. A documentação ainda cita a URL dp-corporate-frontend.",
        "Movida para a aba Discutir com a equipe.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/470121/details",
        32,
    ),
    (
        CARD,
        "Encerrada pela equipe",
        "451674",
        "[DP] Migração GCP | Converter ETL's (DataBricks)",
        "Cancelado",
        "",
        "",
        "",
        "",
        "Complexidade Cancelado.",
        "Fora da ordem.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451674/details",
        38,
    ),
    (
        CARD,
        "Encerrada pela equipe",
        "424941",
        "[DP] PostgreSQL - Levantamento de recursos para reescrita de código",
        "Cancelado",
        "",
        "",
        "",
        "",
        "Complexidade Cancelado.",
        "Fora da ordem.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/424941/details",
        42,
    ),
    (
        CARD,
        "Encerrada pela equipe",
        "451676",
        "[DP] Migração GCP | Migrar banco de dados SQL Server -- ABORTADO",
        "Concluído",
        "",
        "",
        "",
        "",
        "Título ABORTADO e status Concluído. A documentação diz que a réplica 3 não será migrada e está desligada. Master com 10 vCores, réplica com 18, HML com 6, staging com 2. Tamanho citado: 461 GB.",
        "Fora do caminho. A existência de réplica em produção não virou atividade.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/451676/details",
        37,
    ),
    (
        CARD,
        "Encerrada pela equipe",
        "476063",
        "[DATP] DasaLogs -> api local - unidades",
        "Stand By",
        "",
        "",
        "",
        "",
        "Complexidade Stand By.",
        "Fora da ordem.",
        "https://dasa.businessmap.io/ctrl_board/96/cards/476063/details",
        33,
    ),
    (
        GAP_DOC,
        "2. HML — antes de subir workload",
        "",
        "Autenticação dos canais na HML: API key, Azure AD e SSO",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        0,
    ),
    (
        GAP_DOC,
        "3. HML — APIs e gateway, em paralelo com o Airflow",
        "",
        "dp-api-gtw e dp-cad-proxy do diagrama",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        0,
    ),
    (
        GAP_DOC,
        "3. HML — APIs e gateway, em paralelo com o Airflow",
        "",
        "API de agendamento documentada (busca-agendamento-backend)",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        0,
    ),
    (
        GAP_DOC,
        "6. HML — dado, depois do Airflow e do backend",
        "",
        "Conectividade com as fontes do diagrama",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        0,
    ),
    (
        GAP_DOC,
        "6. HML — dado, depois do Airflow e do backend",
        "",
        "Staging SQL no caminho do Airflow em HML",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        0,
    ),
]

REGRAS = [
    "A ordem segue dependência e arquitetura. A coluna Cronograma original guarda a data da planilha e não foi usada para ordenar. Várias datas se contradizem: portal e Blob em fevereiro de 2027, adequação do Airflow em março–abril de 2027, Airflow com nome prd em 09/10/26.",
    "O teste funcional (card 451689) e o teste de carga fecham o que a HML precisa para testar. A linha de Databricks na GCP fica depois desse teste e antes da produção, e não bloqueia o teste. O único cluster concluído é o GKE de HML.",
    "A passagem do Data Factory para o Airflow não entra nesta aba. Está concluída e é pré-requisito de outro projeto. O Airflow de partida está na Azure. Nesta lista a frente é migrar esse Airflow para a GCP: primeiro a spike, por causa da versão 2.2.5, depois a adequação. Carga, Blob e portal esperam essa migração. As APIs que não citam Airflow nem Cognitive Search não esperam.",
    "Exam fica depois do índice. As outras APIs não foram colocadas atrás do índice: só a exam e o dp-backend-api cancelado citam Cognitive Search. O diagrama desenha busca-search na frente do conjunto de APIs, e essa diferença está na aba Dúvidas.",
    "Lacuna só entra quando a documentação nomeia o componente ou quando um card cita a dependência e não existe card para ela. Lacuna não tem Card ID, responsável, complexidade nem prazo.",
    "A coluna Origem fica 'IA' quando a linha não tinha card. A base explica o que se espera e a coluna ao lado das observações diz por que a linha está naquele ponto. Essas linhas não têm responsável, prazo nem Card ID. Origem 'Card da planilha' é atividade da equipe.",
    "O card de BigQuery no lugar do Databricks, o DasaLogs, o frontend dp-corporate-frontend e o repositório datp-backoffice-frontend-portal foram para a aba Discutir com a equipe. São itens da planilha que não se sustentam neste caminho.",
    "Cards cancelados ou abortados que não entram nessa discussão ficam no fim, na etapa Encerrada pela equipe. O status deles sai da coluna Complexidade.",
]

DUVIDAS = [
    (
        "1",
        "O banco criado no card 451661 já tem, em HML, exame, unidade e convênio?",
        "Se já tiver, a lacuna de carga sai da frente do teste funcional. Se não tiver, ela continua depois da migração do Airflow da Azure para a GCP. A migração do SQL Server está abortada.",
    ),
    (
        "2",
        "Quais APIs leem o Cognitive Search além da exam?",
        "O diagrama coloca busca-search na frente das APIs. A planilha só trava a exam e o dp-backend-api cancelado. Se brand, local, health insurance, elegibility ou commons leem o índice, elas saem da etapa 3 e vão para depois do card 451682.",
    ),
    (
        "3",
        "O que o Databricks de HML entrega para a API ou para o portal?",
        "A linha dos workspaces foi inclusa por gap e ficou depois do teste funcional, porque o diagrama até as APIs não passa pelo Databricks. Se a API ou o portal consomem a saída, ela sobe para antes do card 451689.",
    ),
    (
        "4",
        "datp-searcher-api-elegibility é a API de agendamento do documento?",
        "O documento lista agendamento entre as seis APIs. O diagrama nomeia busca-agendamento-backend. A planilha tem elegibility e não tem card chamado agendamento. Se forem duas APIs, falta o workload de agendamento.",
    ),
    (
        "5",
        "O card cancelado dp-corporate-frontend (470121) é o mesmo frontend do card 469051?",
        "A documentação cita a URL dp-corporate-frontend e o repositório datp-backoffice-frontend. Um card está cancelado e o outro está na fila de HML. Sem essa resposta os dois não foram unidos.",
    ),
    (
        "6",
        "O download D+1 de cadastro lê o Blob?",
        "O Blob depende de Airflow, backend e frontend. O portal depende de Airflow, cargas de cadastro e D+1. A planilha não liga um ao outro. Se o D+1 ler o Blob, o card 451688 fica antes das cargas.",
    ),
    (
        "7",
        "A conectividade do card 424949 inclui as fontes do diagrama?",
        "K2/Gliese, base federada, MDM, GSC, hospitais e MV estão no desenho. A seção de rede do documento está vazia. O texto dos workloads fala só do IP da aplicação. Se as fontes forem outro trabalho, falta uma atividade antes da carga.",
    ),
    (
        "8",
        "No teste de HML a autenticação continua como está?",
        "O documento marca Azure AD como dependência interna que não pode ser desacoplada. Frontend do DP usa AD, portal usa SSO, APIs usam API key. O template GCP escreve Azure AD para Google IAM e não foi preenchido. Nenhuma atividade de identidade foi criada.",
    ),
    (
        "9",
        "O que entra no teste de carga?",
        "Não há card, volume, API nem meta. O template do documento cita Performance sem conteúdo. A linha está depois do teste funcional e antes da produção, sem cenário.",
    ),
    (
        "10",
        "dp-api-gtw e dp-cad-proxy entram nesta migração?",
        "Estão no diagrama AS IS e não têm card. Ficaram fora do caminho. O Zuul tem card próprio.",
    ),
    (
        "11",
        "O staging SQL ainda faz parte do caminho em HML?",
        "STAGING_PRD está na documentação e não há card. Ficou fora do caminho.",
    ),
    (
        "12",
        "Event Hub entra no teste funcional de algum canal?",
        "O card manda deixar por último e nenhuma API o cita. Ele está no fim da produção. Se NAV ou CDC dependem do evento para o teste de HML, ele sobe para antes do card 451689.",
    ),
]


def etapa_color(etapa):
    n = etapa.split(".", 1)[0]
    return {
        "1": "E7E6E6",
        "2": "D6E3F0",
        "3": "DEEBF7",
        "4": "FFF2CC",
        "5": "F8CBAD",
        "6": "FCE4D6",
        "7": "FCE4D6",
        "8": "C6EFCE",
        "9": "E2D5F1",
        "10": "F4B183",
    }.get(n, "F2F2F2")


SEQUENCE = [
    "424950",
    "451661",
    "451684",
    "476416",
    "424949",
    "Autenticação dos canais na HML: API key, Azure AD e SSO",
    "424942",
    "470110",
    "470109",
    "dp-api-gtw e dp-cad-proxy do diagrama",
    "470118",
    "470115",
    "470113",
    "API de agendamento documentada (busca-agendamento-backend)",
    "470112",
    "424944",
    "469049",
    "469051",
    "451682",
    "470114",
    "478326",
    "Criar historia adequação airflow, vinculada ao fim da poc do Victor => necessario devido versao de airflow",
    "Conectividade com as fontes do diagrama",
    "Staging SQL no caminho do Airflow em HML",
    "Carga, no banco de HML, dos dados de exame, unidade e convênio que hoje estão no Corporate",
    "451688",
    "Cargas de cadastro",
    "Downloads D+1 de cadastro",
    "468966",
    "470165",
    "470100",
    "451681",
    "451677",
    "451689",
    "Teste de carga em HML",
    "Workspaces Databricks de HML e de PRD na GCP",
    "GKE de produção",
    "470150",
    "470149",
    "470148",
    "470146",
    "470145",
    "468978",
    "451683",
    "451691",
    "470119",
    "451674",
    "424941",
    "451676",
]


def sequence_key(row):
    card_id, descricao = row[2], row[3]
    token = card_id or descricao
    try:
        return SEQUENCE.index(token)
    except ValueError as exc:
        raise SystemExit(f"fora da sequência: {token}") from exc


def origem_font(origem):
    if origem == CARD:
        return Font(name="Calibri", size=10)
    return Font(name="Calibri", size=10, italic=True, color="833C0C")


def header_style():
    return {
        "fill": PatternFill("solid", fgColor="1F4E79"),
        "font": Font(color="FFFFFF", bold=True, name="Calibri", size=11),
        "border": Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        ),
    }


def prepare_sheet(ws, last_col, last_row):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{last_row}"
    ws.auto_filter.add_sort_condition("A2:A200")
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.6
    ws.print_title_rows = "1:1"
    ws.oddFooter.right.text = "Página &P de &N"
    ws.sheet_view.zoomScale = 110
    ws.page_setup.scale = 70


def write_headers(ws, headers):
    style = header_style()
    for col, text in enumerate(headers, 1):
        cell = ws.cell(1, col, text)
        cell.fill = style["fill"]
        cell.font = style["font"]
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = style["border"]
    ws.row_dimensions[1].height = 32
    return style["border"]


DISCUSS_IDS = ("478328", "476063", "470121", "468985")

DISCUSS_WHY = {
    "478328": (
        "A equipe pediu levar a orquestração do Databricks para o BigQuery. "
        "O diagrama AS IS não passa pelo Databricks até as APIs. A ida para a GCP "
        "não exige essa troca: a decisão deste plano é levar os workspaces, e o "
        "card de converter os ETLs já está cancelado. O desenho com bronze, silver "
        "e gold está no documento sem a estratégia preenchida."
    ),
    "476063": (
        "A equipe deixou em stand-by. O documento de arquitetura não cita DasaLogs "
        "no caminho de exame, unidade e convênio, nem nas APIs, no Airflow ou na busca."
    ),
    "470121": (
        "A equipe cancelou o workload dp-corporate-frontend. A documentação ainda "
        "aponta essa URL como frontend do DP. O card 469051 segue no caminho com outro nome. "
        "Os dois não foram tratados como o mesmo frontend."
    ),
    "468985": (
        "A equipe cancelou o workload do repositório datp-backoffice-frontend-portal. "
        "A documentação cita esse repositório como frontend do portal de cadastro. "
        "Os cards 470165 e 451677 seguem no caminho com outros nomes."
    ),
}

STATUS_DO_CARD = {
    "Concluído",
    "Cancelado",
    "Stand By",
    "Sem Definição de Escopo",
}
COMPLEXIDADE = {"Alta", "Média", "Baixa"}


def separar_status_e_complexidade(valor):
    texto = (valor or "").strip()
    if not texto:
        return None, None
    if texto in STATUS_DO_CARD:
        return texto, None
    if texto in COMPLEXIDADE:
        return None, texto
    raise SystemExit(f"valor inesperado na coluna Complexidade: {texto}")


def origem_visivel(origem):
    if origem != CARD:
        return "IA"
    return origem


STORIES = {
    "424950": (
        "Homologação já tem cluster GKE. É aqui que as APIs sobem para o NAV e o CDC experimentarem antes da produção.",
        "Começo do caminho. Sem este cluster não há onde publicar o DP.",
    ),
    "451661": (
        "O banco de homologação já foi criado. A migração do SQL Server foi abortada, então este banco é o que as APIs vão ler no lugar do Corporate.",
        "Já concluído. A carga dos dados entra mais adiante, senão a API sobe vazia.",
    ),
    "451684": (
        "O Redis já foi migrado. No diagrama ele é o busca-cache, entre o Corporate e as APIs. O teste deste cache espera os pods no ar.",
        "Concluído na base. A prova dele acontece quando as APIs estiverem de pé.",
    ),
    "476416": (
        "A continuação do workload da API commons já foi concluída. O card 470112, da criação, ainda está em andamento.",
        "Fica na base concluída. A criação restante segue com as outras APIs.",
    ),
    "424949": (
        "Os workloads registram que a conectividade por IP está errada. O esperado é um endereço de homologação que o NAV e o CDC consigam chamar.",
        "Antes de subir as APIs. Se o canal continuar no IP antigo, a HML não é usada.",
    ),
    "Autenticação dos canais na HML: API key, Azure AD e SSO": (
        "O documento diz que o Azure AD não pode ser desacoplado. As APIs usam API key, o frontend do DP usa AD e o portal usa SSO. O esperado é a HML aceitar esse mesmo acesso.",
        "Antes das APIs. O canal não deve ganhar outra identidade para chamar a homologação.",
    ),
    "424942": (
        "A esteira precisa publicar no GKE de homologação. Nenhum workload cita este card como bloqueio.",
        "Em paralelo com a conectividade. Não segura a subida do primeiro workload.",
    ),
    "470110": (
        "API de brand, uma das consultas do documento. O esperado é ela responder em HML sem depender de Airflow nem de busca.",
        "Sobe com as APIs que o canal chama direto, depois da conectividade e da autenticação.",
    ),
    "470109": (
        "Gateway Zuul. No diagrama ele fica na entrada, com autorização, limite de taxa e WAF. O esperado é o NAV e o CDC entrarem por aqui.",
        "Junto das APIs. É a porta de entrada da homologação.",
    ),
    "dp-api-gtw e dp-cad-proxy do diagrama": (
        "O diagrama AS IS nomeia dp-api-gtw e dp-cad-proxy e não há card. O esperado é decidir se o Zuul já cobre os dois ou se cada um precisa de workload em HML.",
        "Ao lado do gateway. Sem essa decisão não se sabe se o canal perde uma porta.",
    ),
    "470118": (
        "API de local, a consulta de unidade. O esperado é o canal consultar unidade na HML, sem Airflow e sem índice de busca.",
        "Na mesma fila das APIs de consulta direta.",
    ),
    "470115": (
        "API de convênio. O esperado é o canal consultar convênio na HML.",
        "Na mesma fila das APIs de consulta direta.",
    ),
    "470113": (
        "Workload de elegibility. O documento chama a outra API de agendamento e o diagrama mostra busca-agendamento-backend. Este card não usa esse nome.",
        "Sobe com as APIs que não citam busca. A linha de agendamento, ao lado, cobre o nome da documentação.",
    ),
    "API de agendamento documentada (busca-agendamento-backend)": (
        "O documento lista agendamento entre as seis APIs e o diagrama mostra busca-agendamento-backend. Não há card com esse nome. O esperado é essa consulta existir em HML. Se for o mesmo serviço da elegibility (470113), as duas linhas viram uma.",
        "Na fila das APIs, para o nome do documento não ficar sem dono.",
    ),
    "470112": (
        "API commons. A continuação 476416 está concluída e este card segue em andamento. O esperado é ela no ar com as outras.",
        "Na mesma fila, antes do teste de conectividade.",
    ),
    "424944": (
        "Teste de conectividade das APIs, não teste funcional. O esperado é confirmar que a aplicação enxerga o destino novo, ainda sem exigir o dado completo.",
        "Logo depois dos workloads desta fila. O teste com o canal espera o dado e a busca.",
    ),
    "469049": (
        "Backend do frontend do DP, repositório datp-backoffice-backend. O esperado é o backoffice responder em HML. Não cita Airflow.",
        "Sobe com as APIs. O teste da tela espera a integração.",
    ),
    "469051": (
        "Frontend do DP. A documentação cita a URL dp-corporate-frontend; o workload com esse nome está cancelado e ficou para discussão. O esperado é esta adaptação ser a tela aberta em HML, testada só com tudo integrado.",
        "A adaptação entra aqui. O teste da tela é o card 451689.",
    ),
    "451682": (
        "O diagrama coloca o busca-search entre o Corporate e as APIs. O documento marca o Cognitive Search como dependência alta, ainda sem decisão de desacoplamento. O card troca o índice para Vertex. O esperado é a busca de exam funcionar como hoje.",
        "Antes da API de exam. As outras APIs não esperam este card: a planilha só cita a busca na exam.",
    ),
    "470114": (
        "API de exam. Depende do Cognitive Search e também registra conectividade por IP errada. O esperado é o canal buscar exame na HML pelo índice novo.",
        "Depois do índice e depois da conectividade.",
    ),
    "478326": (
        "O Airflow já está na Azure. Esta spike, em docker local, acha o que quebra da versão 2.2.5 para a versão do Google. O esperado é a lista de incompatibilidades antes de migrar.",
        "Abre a frente do Airflow, em paralelo com as APIs. A migração e a carga esperam o resultado.",
    ),
    "Criar historia adequação airflow, vinculada ao fim da poc do Victor => necessario devido versao de airflow": (
        "É a ida do Airflow da Azure para a GCP, por causa da versão. Não tem Card ID. Portal, Blob e backoffice dependem desse Airflow já na GCP. Os workloads com prd no nome ficam na produção.",
        "Depois da spike. Carga e portal vêm em seguida.",
    ),
    "Conectividade com as fontes do diagrama": (
        "O diagrama lê K2/Gliese, base federada, MDM, GSC, hospitais e MV. O card 424949 não cita essas fontes. O esperado é o Airflow na GCP alcançar as mesmas origens que hoje alimentam o Corporate.",
        "Depois do Airflow na GCP e antes da carga. Sem as fontes o banco de HML não recebe dado.",
    ),
    "Staging SQL no caminho do Airflow em HML": (
        "A documentação lista o banco STAGING_PRD, 2 vCores, no caminho até o Corporate. O esperado é saber se o Airflow que vai à GCP ainda grava nesse staging. Se não gravar, a linha sai.",
        "Junto da migração do Airflow, antes da carga do banco que as APIs leem.",
    ),
    "Carga, no banco de HML, dos dados de exame, unidade e convênio que hoje estão no Corporate": (
        "As APIs servem exame, unidade e convênio. O banco de HML existe e a cópia do SQL Server foi abortada. O esperado é esse banco receber os dados do Corporate, para o NAV e o CDC não consultarem vazio.",
        "Depois do Airflow na GCP e das fontes. Antes do teste com os canais.",
    ),
    "451688": (
        "Há 269 GB de Blob. A adaptação depende de Airflow, backend e frontend. O esperado é a aplicação ler o arquivo na GCP.",
        "Depois do Airflow e do backend e do frontend do DP. A planilha não diz se o download D+1 usa este Blob.",
    ),
    "Cargas de cadastro": (
        "Os cards 470165 e 470100 dependem das cargas de cadastro e não existe card só para elas. O diagrama tem o Portal de Cadastros. O esperado é a carga rodar no Airflow da GCP antes do portal abrir.",
        "Antes dos workloads de portal que citam essa dependência.",
    ),
    "Downloads D+1 de cadastro": (
        "Os mesmos cards dependem do download D+1 de cadastro. Não há card só para isso. O esperado é o arquivo do dia seguinte disponível em HML.",
        "No mesmo bloco das cargas de cadastro, antes do portal.",
    ),
    "468966": (
        "Backend do backoffice. A observação diz que depende do Airflow. O esperado é ele no ar depois que o Airflow estiver na GCP.",
        "Abre o bloco do portal, antes do teste integrado.",
    ),
    "470165": (
        "Tela portal-cadastro-web. Depende do Airflow pronto, das cargas de cadastro e do D+1. O repositório datp-backoffice-frontend-portal está cancelado e em discussão. O esperado é o portal de cadastro abrir em HML com esses dados.",
        "Depois do Airflow, das cargas e do D+1.",
    ),
    "470100": (
        "Frontend do backoffice, com a mesma dependência do portal-cadastro-web. O esperado é a mesma condição: Airflow, cargas e D+1 prontos.",
        "Junto do portal-cadastro-web.",
    ),
    "451681": (
        "Adaptação do backend do portal de cadastro. O esperado é o código ajustado para a HML, no mesmo bloco do portal.",
        "Com os workloads do portal. A data de fevereiro da planilha não define a ordem.",
    ),
    "451677": (
        "Adaptação do frontend do portal de cadastro. Só testa com tudo integrado. O esperado é a tela pronta para o teste do card 451689.",
        "A adaptação fica neste bloco. O teste da tela fecha a HML.",
    ),
    "451689": (
        "Único card de teste da equipe, sem escopo. O documento não define SLA. O esperado é o NAV, o CDC e o portal usarem exame, unidade e convênio na HML, com o mesmo contrato de hoje, e haver plano de rollback.",
        "Fecha a homologação, depois das APIs, da busca, do dado e do portal.",
    ),
    "Teste de carga em HML": (
        "Não há card, volume nem meta. O documento cita performance sem plano. O esperado é exercitar a HML depois do teste funcional e antes de planejar produção.",
        "Logo depois do teste funcional. Produção só começa depois.",
    ),
    "Workspaces Databricks de HML e de PRD na GCP": (
        "A decisão oficial é levar o Databricks para a GCP, não o desenho de BigQuery. A documentação lista o workspace de PRD adb-7665073515379548 e o de HML adb-887971986833958, sem jobs. O esperado é os dois workspaces rodarem na GCP. O card 478328 ficou na aba de discussão.",
        "Depois do teste dos canais e antes da produção. O diagrama até as APIs não passa pelo Databricks, então o NAV e o CDC não esperam esta linha.",
    ),
    "GKE de produção": (
        "A documentação cita o cluster aks-prd-cadastro01. O único GKE com card é o de HML, já concluído. O esperado é o cluster de produção existir só depois da HML testada.",
        "Primeira linha de produção.",
    ),
    "470150": (
        "Worker do Airflow de produção. O esperado é esta peça no ar depois da HML. A data 09/10/26 da planilha não antecipa a linha.",
        "No bloco de produção, na ordem original da planilha.",
    ),
    "470149": (
        "Workload principal do Airflow de produção. O esperado é ele no ar depois da HML.",
        "No bloco de produção, na ordem original da planilha.",
    ),
    "470148": (
        "Sincronização de usuários do Airflow de produção. O esperado é ela no ar depois da HML.",
        "No bloco de produção, na ordem original da planilha.",
    ),
    "470146": (
        "Scheduler do Airflow de produção. O esperado é ele no ar depois da HML.",
        "No bloco de produção, na ordem original da planilha.",
    ),
    "470145": (
        "Flower do Airflow de produção. O esperado é ele no ar depois da HML.",
        "No bloco de produção, na ordem original da planilha.",
    ),
    "468978": (
        "Migração do banco do Airflow de produção. O esperado é ela rodar depois da HML.",
        "No bloco de produção, na ordem original da planilha.",
    ),
    "451683": (
        "A documentação lista o Event Hub data-provider-prd. A planilha manda deixar por último e nenhuma API cita esse evento. O esperado é a troca para Pub/Sub na produção, não na prova dos canais.",
        "No fim da produção, antes do cutover.",
    ),
    "451691": (
        "Go-live sem escopo. O documento não define se a entrada é única ou gradual, nem DNS nem janela. O esperado é o plano de produção, só depois da HML aceita pelos canais.",
        "Última linha de produção.",
    ),
    "470119": (
        "Cancelado pela equipe, com observação de Cognitive Search. Os backends 469049 e 468966 seguem no caminho. O documento não descreve este card além disso.",
        "Fora do caminho aberto. O cancelamento não tira a API de exam, que continua atrás do índice.",
    ),
    "451674": (
        "Cancelado. A decisão oficial é levar o Databricks, não converter os ETLs para outra ferramenta.",
        "Encerrado. Não volta para o caminho.",
    ),
    "424941": (
        "Cancelado. Não há reescrita para PostgreSQL neste plano. O banco de HML já foi criado.",
        "Encerrado.",
    ),
    "451676": (
        "Abortado e concluído. O SQL Server não é copiado. A documentação diz que a réplica 3 não migra e está desligada. O dado segue pela carga no banco já criado.",
        "Encerrado. A carga de HML cobre o que as APIs precisam ler.",
    ),
}


def write_main(wb):
    ws = wb.active
    ws.title = "BM - GCP Atividades - DP"
    headers = [
        "Ordem",
        "Etapa",
        "Descrição",
        "Origem",
        "Card ID",
        "Status",
        "Complexidade",
        "Depêndencia",
        "Cronograma original (não usado na ordem)",
        "Recurso Responsável",
        "Observações",
        "Por que nesta posição",
        "Base na planilha ou na documentação",
        "Card URL",
        "Linha na planilha original",
    ]
    border = write_headers(ws, headers)
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    ai_font = Font(name="Calibri", size=10, bold=True, color="843C0C")
    ordered = sorted(
        (row for row in ROWS if row[2] not in DISCUSS_IDS),
        key=sequence_key,
    )
    for i, row in enumerate(ordered, 1):
        (
            origem,
            etapa,
            card_id,
            descricao,
            complexidade,
            dependencia,
            cronograma,
            responsavel,
            observacoes,
            base,
            posicao,
            url,
            linha,
        ) = row
        historia = STORIES.get(card_id or descricao)
        if historia:
            base, posicao = historia
        status, complexidade_somente = separar_status_e_complexidade(complexidade)
        origem_coluna = origem_visivel(origem)
        values = [
            i,
            etapa,
            descricao,
            origem_coluna,
            card_id or None,
            status,
            complexidade_somente,
            dependencia or None,
            cronograma or None,
            responsavel or None,
            observacoes or None,
            posicao,
            base,
            url or None,
            linha or None,
        ]
        fill = PatternFill("solid", fgColor=etapa_color(etapa))
        font = origem_font(origem)
        for col, value in enumerate(values, 1):
            cell = ws.cell(i + 1, col, value)
            cell.font = ai_font if col == 4 and origem != CARD else font
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border
        if url:
            url_cell = ws.cell(i + 1, 14)
            url_cell.hyperlink = url
            url_cell.font = link_font
        ws.row_dimensions[i + 1].height = 52
    widths = [8, 46, 62, 42, 12, 24, 16, 46, 28, 26, 36, 55, 55, 40, 14]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    prepare_sheet(ws, len(headers), ws.max_row)


def write_discuss(wb):
    ws = wb.create_sheet("Discutir com a equipe", 1)
    headers = [
        "Ordem",
        "Descrição",
        "Card ID",
        "Status",
        "Complexidade",
        "Depêndencia",
        "Cronograma original",
        "Recurso Responsável",
        "Observações",
        "Por que não entra neste caminho",
        "Card URL",
        "Linha na planilha original",
    ]
    border = write_headers(ws, headers)
    by_id = {row[2]: row for row in ROWS if row[2] in DISCUSS_IDS}
    fill = PatternFill("solid", fgColor="F8D7DA")
    font = Font(name="Calibri", size=10)
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    for i, card_id in enumerate(DISCUSS_IDS, 1):
        (
            _origem,
            _etapa,
            _card_id,
            descricao,
            complexidade,
            dependencia,
            cronograma,
            responsavel,
            observacoes,
            _base,
            _posicao,
            url,
            linha,
        ) = by_id[card_id]
        status, complexidade_somente = separar_status_e_complexidade(complexidade)
        values = [
            i,
            descricao,
            card_id,
            status,
            complexidade_somente,
            dependencia or None,
            cronograma or None,
            responsavel or None,
            observacoes or None,
            DISCUSS_WHY[card_id],
            url or None,
            linha or None,
        ]
        for col, value in enumerate(values, 1):
            cell = ws.cell(i + 1, col, value)
            cell.font = font
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border
        if url:
            url_cell = ws.cell(i + 1, 11)
            url_cell.hyperlink = url
            url_cell.font = link_font
        ws.row_dimensions[i + 1].height = 60
    widths = [8, 62, 12, 22, 16, 42, 24, 24, 36, 70, 42, 14]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    prepare_sheet(ws, len(headers), ws.max_row)
    ws.oddHeader.left.text = "Discutir com a equipe"
    ws.page_setup.orientation = "landscape"
    ws.sheet_view.showGridLines = False


def write_rules(wb):
    ws = wb.create_sheet("Regra de ordem")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Como a aba foi ordenada"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    for i, text in enumerate(REGRAS, 3):
        cell = ws.cell(i, 1, text)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.font = Font(name="Calibri", size=11)
        ws.row_dimensions[i].height = 48
    ws.column_dimensions["A"].width = 150
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.6
    ws.print_title_rows = "1:1"
    ws.oddFooter.right.text = "Página &P de &N"


def write_duvidas(wb):
    ws = wb.create_sheet("Dúvidas")
    headers = ["#", "Dúvida", "Como a resposta mexe na ordem"]
    border = write_headers(ws, headers)
    fill = PatternFill("solid", fgColor="FFF8E7")
    font = Font(name="Calibri", size=10)
    for i, item in enumerate(DUVIDAS, 2):
        for col, value in enumerate(item, 1):
            cell = ws.cell(i, col, value)
            cell.font = font
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border
        ws.row_dimensions[i].height = 48
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 78
    ws.column_dimensions["C"].width = 88
    prepare_sheet(ws, 3, ws.max_row)


def main():
    wb = Workbook()
    write_main(wb)
    write_discuss(wb)
    write_rules(wb)
    write_duvidas(wb)
    wb.save(OUT)


if __name__ == "__main__":
    main()
