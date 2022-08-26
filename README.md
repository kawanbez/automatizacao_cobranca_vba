![background](https://github.com/kawanbez/automatizacao_cobranca_vba/blob/main/bg2.png)

# [PORTUGUÊS] PT-BR (English version bellow)
# Automatização do Processo de Cobrança - VBA & MySQL

Inicialmente o processo de cobrança na empresa em que atuo era totalmente manual, conforme o seguinte processo:

* Download de um arquivo com os pedidos criados em D-1 com informações pertinentes a área financeira, como número do pedido, responsável pela venda, data de vencimento do boleto, data de criação do pedido, etc
* Esses dados eram consolidados diariamente para que com base nas informações dessa tabela, pudesse ser criada uma "tabela filha" com informações pertinentes a área de cobrança. Em resumo, era criado uma nova tabela com base na data de vencimento dos pedidos e a partir do momento do vencimento do pedido, era realizado a cobrança de forma totalmente manual, desde a localização do cliente que seria cobrado como até a escrita do email pelo atendente da cobrança.

Esse processo muitas vezes era suscetível a erros, pois muitas vezes, a planilha pai não era atualizada com prorrogações, cancelamentos e suspensões, gerando cobranças indevidas.

### Localizando as possíveis falhas e dificuldades no processo

O primeiro passo para a correção e instauração de um novo processo de cobrança era entender os gaps processuais e as dificuldades da equipe responsável pela ação.
Haviam falhas processuais tanto no controle de pedidos, na atualização dessas informações como no processo de cobrança, onde os pedidos cobrados muitas vezes caiam em um limbo caso o atendente não se atentasse a seu email.

##### Planilha de Controle de Pedidos

A planilha de controle de pedidos era o ponto inicial para a correção do processo. A base crua era disponibilizada pela plataforma de BI, entretanto, essa base era atualizada entre 1h e 2h da manhã, todo dia, com dados referentes ao dia anterior. Devido ao fato dos dados estarem sempre em D-1, muitas vezes havia gap de informações pois alterações feitas após a atualização dos dados impactavam diretamente no fluxo de inserção desses pedidos na base. Portanto, a primeira mudança que realizei foi nesse processo. 
Solicitei que parassem de utilizar dados em D-1 e começassem a utilizar dados em tempo real com auxilio da query que desenvolvi em Mysql:

~~~~sql

SELECT o.id
  , o.customer_id
  , o.status_id
  , IFNULL(i.number, 'NOT ISSUED')         AS 'invoice id'
  , IFNULL(i.type, '') AS 'Invoice Type'
  , CONVERT(ROUND(o.total_price, 2), CHAR) AS 'order total price'
  , opt.status
  , CONVERT(MAX(opt.updated_at), DATE) AS 'Updated Date'
  , IFNULL(opt.expiry_date, '')            AS 'expiry date'
  , c.email                                AS 'customer email'
  , c.type                                 AS 'customer subtype'
  , IFNULL(c.company_name, '')             AS 'company name'
  , opt.method
  , IFNULL(opt.paid_at, '')                AS 'paid_at'
  , c.service_type
  , IFNULL(u.name, '')
  , IFNULL(u.email,'')                     AS 'sales email'
  , CONVERT(o.created_at, DATE)            AS 'order created at'
  , i.file_url                             AS 'URL NF'
FROM banco_de_dados                          AS o
LEFT JOIN order_invoice                    AS i   ON o.id = i.order_id
LEFT JOIN customer                        AS c   ON o.customer_id = c.id
LEFT JOIN order_payment_transaction       AS opt ON o.id = opt.order_id
LEFT JOIN alpha.user                       AS u   ON o.sales_user_id = u.id
WHERE o.created_at > 20201101
AND opt.status IN ('paid', 'waiting_payment')
AND opt.method = 'invoice'
AND i.type = 'NFSE'
GROUP BY o.id

UNION

SELECT o.id
  , o.customer_id
  , o.status_id
  , IFNULL(i.number, 'NOT ISSUED')         AS 'invoice id'
  , IFNULL(i.type, '') AS 'Invoice Type'
  , CONVERT(ROUND(o.total_price, 2), CHAR) AS 'order total price'
  , opt.status
  , CONVERT(MAX(opt.updated_at), DATE) AS 'Updated Date'
  , IFNULL(opt.expiry_date, '')            AS 'expiry date'
  , c.email                                AS 'customer email'
  , c.type                                 AS 'customer subtype'
  , IFNULL(c.company_name, '')             AS 'company name'
  , opt.method
  , IFNULL(opt.paid_at, '')                AS 'paid_at'
  , c.service_type
  , IFNULL(u.name,'')
  , IFNULL(u.email,'')                     AS 'sales email'
  , CONVERT(o.created_at, DATE)            AS 'order created at'
  , i.file_url                             AS 'URL NF'
FROM banco_de_dados                           AS o
LEFT JOIN order_invoice                    AS i   ON o.id = i.order_id
LEFT JOIN customer                        AS c   ON o.customer_id = c.id
LEFT JOIN order_payment_transaction       AS opt ON o.id = opt.order_id
LEFT JOIN alpha.user                       AS u   ON o.sales_user_id = u.id
WHERE o.created_at > 20201101
and opt.status = 'cancelled'
and opt.method = 'invoice'
AND i.type = 'NFSE'
and o.id NOT IN (SELECT t.id 
                 FROM alpha.order t 
                 JOIN alpha.order_payment_transaction pt ON t.id = pt.order_id
                 AND pt.status != 'cancelled')
GROUP BY o.id 
~~~~

Com base nesse código, temos as mesmas informações que a plataforma de BI, só que sempre em tempo real, portanto, as informações que plotariamos na tabela sempre estariam atualizadas. 
Uma vez que a base estava com dados confiáveis, era o momento de mudar o processo de cobrança

#### Estabelecer uma régua de cobrança

Como dito anteriormente, a cobrança era feita de forma manual e dependia totalmente do controle de emails de quem realizou a cobrança, dificultando outras pessoas da area localizarem históricos e acompanharem o dia a dia do time responsável pela ação. Portanto, era necessário criar uma rotina de cobrança para facilitar a localização de informações e checagem do andamento dos pagamentos em atraso.
Devido a isso, a régua de cobrança abaixo foi desenvolvida:

![background](https://github.com/kawanbez/automatizacao_cobranca_vba/blob/main/r%C3%A9gua%20de%20cobran%C3%A7a.png)

Agora a cobrança contempla:
* Um lembrete aos vendedores responsáveis pelos pedidos à vencer em até 5 dias
* Um lembrete de vencimento para os clientes com pedidos à vencer em 2 dias
* Uma ação de retirada da possibilidade de compra faturada pelos clientes com pedidos vencidos em D-1
* A primeira cobrança para clientes com pedidos vencidos em D-1
* O reenvio do boleto e da nota fiscal para clientes com pedidos vencidos em D-3
* A segunda cobrança para clientes com pedidos vencidos em D-3
* Um aviso de inclusão aos órgãos de proteção ao crédito para clientes com pedidos vencidos em D-10
* A inclusão nos órgãos de proteção ao crédito

Após determinarmos o novo fluxo de cobrança, era o momento de automatizar o processo.

#### Criação do código em VBA 

Para a criação do código, se faz necessário uma planilha de apoio oriunda da régua de cobrança, com abas com os status da cobrança: Primeira, segunda e terceira cobrança.
Com a planilha de apoio criada, ainda havia o esforço de enviar os emails, e visando corrigir essa dificuldade, o código em VBA a seguir foi inserido na planilha de apoio:

´´´´

    Sub enviar_email_planilha_1()
    'Cria as variáveis do Outlook
    Dim outapp As Outlook.Application
    Dim outmail As Outlook.MailItem
    Dim numLinha1 As Integer
    
    numLinha1 = 7
    While Planilha2.Cells(numLinha1, 1).Value <> ""
    
    'Cria e chama os objetos do outlook
    Set outapp = CreateObject("outlook.application")
    Set outmail = outapp.CreateItem(olMailItem)
    
    'Desabilita a mensagem de alerta
    Application.DisplayAlerts = False
    
        With outmail
            'email do destinatário
            .To = Planilha2.Cells(numLinha1, 11).Value 'Range("K7")
            'email em cópia
            .CC = Planilha2.Cells(numLinha1, 12).Value 'Range("L7")
            'email em cópia oculta
            .BCC = Planilha2.Cells(numLinha1, 13).Value 'Range("M7")
            'título do email
            .Subject = Planilha2.Cells(numLinha1, 1).Value 'Range("A7")
            'Corpo do email
            .HTMLBody = Planilha2.Cells(numLinha1, 15).Value 'Range("O7")
            .Send
        End With
        
        numLinha1 = numLinha1 + 1
    'resetar os sets
    Set outmail = Nothing
    Set outapp = Nothing
    
    Wend
    Application.DisplayAlerts = True
    
    MsgBox "Cobrança Enviada"
    End Sub
´´´´

Em resumo, com base nas orientações do cabeçalho das abas da planilha, o código faria a leitura e o disparo automatizado dos emails de cobrança pelo Outlook.

### Conclusão

Após a criação do régua de cobrança e automatização de processos, pegamos todos os clientes que ainda não haviam sido transferidos para PDD e assim a equipe financeira e de cobranças conseguiu recuperar até o final de 2019 (Ago/2019 a Dez/2019), pouco mais de 1,9 milhão de reais de forma organizada, visando sempre isentar os clientes de cobranças indevidas e facilitar os processos do time de cobrança com menos trabalho manual e mais assertividade.

# [ENGLISH VERSION]
# Automation of the Billing Process - VBA & MySQL

Initially, the collection process in the company I work for was completely manual, according to the following process:

* Download a file with the orders created in D-1 with information relevant to the financial area, such as the order number, person responsible for the sale, payment slip due date, order creation date, etc.
* These data were consolidated daily so that, based on the information in this table, a "child table" could be created with information relevant to the billing area. In summary, a new table was created based on the due date of the orders and from the moment the order was due, the charge was carried out completely manually, from the location of the customer who would be charged as well as the writing of the email by the billing clerk.

This process was often susceptible to errors, as the parent spreadsheet was often not updated with extensions, cancellations and suspensions, generating undue charges.

### Locating possible failures and difficulties in the process

The first step in correcting and instituting a new collection process was to understand the procedural gaps and difficulties faced by the team responsible for the action.
There were procedural flaws both in order control, in the updating of this information and in the billing process, where charged orders often fell into limbo if the attendant did not pay attention to their email.

##### Order Control Worksheet

The order control worksheet was the starting point for correcting the process. The raw base was made available by the BI platform, however, this base was updated between 1 am and 2 am, every day, with data referring to the previous day. Due to the fact that the data is always in D-1, there was often an information gap as changes made after updating the data directly impacted the flow of insertion of these orders into the base. So the first change I made was in this process.
I asked them to stop using data in D-1 and start using data in real time with the help of the query I developed in Mysql:

~~~~sql

SELECT o.id
  , o.customer_id
  , o.status_id
  , IFNULL(i.number, 'NOT ISSUED')         AS 'invoice id'
  , IFNULL(i.type, '') AS 'Invoice Type'
  , CONVERT(ROUND(o.total_price, 2), CHAR) AS 'order total price'
  , opt.status
  , CONVERT(MAX(opt.updated_at), DATE) AS 'Updated Date'
  , IFNULL(opt.expiry_date, '')            AS 'expiry date'
  , c.email                                AS 'customer email'
  , c.type                                 AS 'customer subtype'
  , IFNULL(c.company_name, '')             AS 'company name'
  , opt.method
  , IFNULL(opt.paid_at, '')                AS 'paid_at'
  , c.service_type
  , IFNULL(u.name, '')
  , IFNULL(u.email,'')                     AS 'sales email'
  , CONVERT(o.created_at, DATE)            AS 'order created at'
  , i.file_url                             AS 'URL NF'
FROM banco_de_dados                          AS o
LEFT JOIN order_invoice                    AS i   ON o.id = i.order_id
LEFT JOIN customer                        AS c   ON o.customer_id = c.id
LEFT JOIN order_payment_transaction       AS opt ON o.id = opt.order_id
LEFT JOIN alpha.user                       AS u   ON o.sales_user_id = u.id
WHERE o.created_at > 20201101
AND opt.status IN ('paid', 'waiting_payment')
AND opt.method = 'invoice'
AND i.type = 'NFSE'
GROUP BY o.id

UNION

SELECT o.id
  , o.customer_id
  , o.status_id
  , IFNULL(i.number, 'NOT ISSUED')         AS 'invoice id'
  , IFNULL(i.type, '') AS 'Invoice Type'
  , CONVERT(ROUND(o.total_price, 2), CHAR) AS 'order total price'
  , opt.status
  , CONVERT(MAX(opt.updated_at), DATE) AS 'Updated Date'
  , IFNULL(opt.expiry_date, '')            AS 'expiry date'
  , c.email                                AS 'customer email'
  , c.type                                 AS 'customer subtype'
  , IFNULL(c.company_name, '')             AS 'company name'
  , opt.method
  , IFNULL(opt.paid_at, '')                AS 'paid_at'
  , c.service_type
  , IFNULL(u.name,'')
  , IFNULL(u.email,'')                     AS 'sales email'
  , CONVERT(o.created_at, DATE)            AS 'order created at'
  , i.file_url                             AS 'URL NF'
FROM banco_de_dados                           AS o
LEFT JOIN order_invoice                    AS i   ON o.id = i.order_id
LEFT JOIN customer                        AS c   ON o.customer_id = c.id
LEFT JOIN order_payment_transaction       AS opt ON o.id = opt.order_id
LEFT JOIN alpha.user                       AS u   ON o.sales_user_id = u.id
WHERE o.created_at > 20201101
and opt.status = 'cancelled'
and opt.method = 'invoice'
AND i.type = 'NFSE'
and o.id NOT IN (SELECT t.id 
                 FROM alpha.order t 
                 JOIN alpha.order_payment_transaction pt ON t.id = pt.order_id
                 AND pt.status != 'cancelled')
GROUP BY o.id 
~~~~

Based on this code, we have the same information as the BI platform, only always in real time, so the information we would plot in the table would always be up to date.
Once the database had reliable data, it was time to change the billing process

#### Establish a billing rule

As previously mentioned, the collection was done manually and totally depended on the control of the emails of the person responsible for the collection, making it difficult for other people in the area to locate histories and monitor the day-to-day activities of the team responsible for the action. Therefore, it was necessary to create a collection routine to facilitate the location of information and check the progress of late payments.
Due to this, the billing rule below was developed:

![background](https://github.com/kawanbez/automatizacao_cobranca_vba/blob/main/r%C3%A9gua%20de%20cobran%C3%A7a.png)

Now the charge includes:
* A reminder to sellers responsible for orders due within 5 days
* An expiration reminder for customers with orders due in 2 days
* An action to withdraw the purchase possibility invoiced by customers with orders overdue on D-1
* The first charge for customers with orders overdue on D-1
* The re-sending of the slip and invoice for customers with orders overdue on D-3
* The second charge for customers with orders overdue on D-3
* A notice of inclusion to credit protection agencies for customers with orders overdue on D-10
* Inclusion in credit protection agencies

After we determined the new billing flow, it was time to automate the process.

#### Code creation in VBA

To create the code, a support worksheet is needed from the billing rule, with tabs with billing status: First, second and third billing.
With the support worksheet created, there was still the effort to send the emails, and in order to correct this difficulty, the following VBA code was inserted into the support worksheet:

´´´´

    Sub enviar_email_planilha_1()
    'Cria as variáveis do Outlook
    Dim outapp As Outlook.Application
    Dim outmail As Outlook.MailItem
    Dim numLinha1 As Integer
    
    numLinha1 = 7
    While Planilha2.Cells(numLinha1, 1).Value <> ""
    
    'Cria e chama os objetos do outlook
    Set outapp = CreateObject("outlook.application")
    Set outmail = outapp.CreateItem(olMailItem)
    
    'Desabilita a mensagem de alerta
    Application.DisplayAlerts = False
    
        With outmail
            'email do destinatário
            .To = Planilha2.Cells(numLinha1, 11).Value 'Range("K7")
            'email em cópia
            .CC = Planilha2.Cells(numLinha1, 12).Value 'Range("L7")
            'email em cópia oculta
            .BCC = Planilha2.Cells(numLinha1, 13).Value 'Range("M7")
            'título do email
            .Subject = Planilha2.Cells(numLinha1, 1).Value 'Range("A7")
            'Corpo do email
            .HTMLBody = Planilha2.Cells(numLinha1, 15).Value 'Range("O7")
            .Send
        End With
        
        numLinha1 = numLinha1 + 1
    'resetar os sets
    Set outmail = Nothing
    Set outapp = Nothing
    
    Wend
    Application.DisplayAlerts = True
    
    MsgBox "Cobrança Enviada"
    End Sub
´´´´

In short, based on the guidance in the header of the spreadsheet tabs, the code would automatically read and trigger the billing emails from Outlook.

### Conclusion

After the creation of the collection rule and process automation, we took all the customers that had not yet been transferred to PDD and thus the financial and collections team managed to recover until the end of 2019 (Aug/2019 to Dec/2019), little more than 1.9 million reais in an organized manner, always aiming to exempt customers from undue charges and facilitate the collection team's processes with less manual work and more assertiveness.
