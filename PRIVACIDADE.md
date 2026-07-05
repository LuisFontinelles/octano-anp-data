# Política de Privacidade — Octano

**Última atualização: 5 de julho de 2026**

O Octano é um aplicativo iOS gratuito, mantido por um desenvolvedor independente, que ajuda motoristas a escolher onde abastecer usando dados públicos oficiais. Esta política explica, em conformidade com a Lei Geral de Proteção de Dados (Lei nº 13.709/2018 — LGPD), quais dados são tratados e como.

**Controlador:** Luis Fontinelles — contato: **[fontinelles.com/contact](https://fontinelles.com/contact)**

## O resumo honesto

O Octano **não tem cadastro, não tem login, não usa publicidade e não vende dados**. O conteúdo que você cria (preços, votos) fica no seu aparelho. Duas informações saem do aparelho: a **localização aproximada**, enviada ao Google para buscar postos próximos, e **dados de uso anônimos**, enviados ao Firebase (Google) para entendermos como o app é usado e melhorá-lo. Compras opcionais de apoio são processadas pela **Apple** — nunca vemos os dados do seu cartão.

## 1. Dados que o app trata

### 1.1 Localização (com sua permissão)
- **Para quê:** encontrar postos próximos, ordenar a lista por distância e traçar rotas.
- **Como:** processada no aparelho. É enviada **apenas** à API do Google Places (busca de postos) e à Apple (MKDirections, cálculo de rota), que atuam como operadoras dessas funções.
- **Base legal:** consentimento (você autoriza no iOS e pode revogar a qualquer momento em Ajustes → Privacidade → Localização).
- O app **não** armazena histórico de localização em servidor do desenvolvedor.

### 1.2 Dados de uso e diagnóstico (analytics)
- **O que:** eventos de navegação (telas visitadas, tempo em cada tela, toques em botões como buscar, traçar rota, abrir posto), modelo do aparelho e versão do iOS, região aproximada (a partir do IP), e um **identificador pseudônimo do app** gerado pelo Firebase. Também coletamos relatórios de **falha (crash)** para corrigir erros.
- **Para quê:** entender como o app é usado, priorizar melhorias e corrigir problemas. **Não** usamos isso para publicidade e **não** coletamos identificador de publicidade (IDFA).
- **Quem processa:** **Firebase (Google LLC)** — Analytics, Crashlytics, Performance. Os dados são tratados sob a política do Google.
- **Base legal:** legítimo interesse na melhoria do serviço; os dados são pseudonimizados e agregados.
- **Notificações:** se você aceitar receber notificações, um token do Firebase Messaging é usado apenas para entregá-las.

### 1.3 Compras no app (contribuições opcionais)
- O app oferece **contribuições únicas** ("gorjetas") para apoiar o projeto. Elas não desbloqueiam recursos.
- O **pagamento é processado pela Apple** (App Store). O desenvolvedor recebe apenas a confirmação da transação — **nunca** o número do cartão ou dados financeiros.

### 1.4 Dados que o app NÃO coleta
Nome, e-mail, telefone, contatos, fotos, dados do cartão e identificador de publicidade (IDFA). Não há login nem cadastro.

## 2. Serviços de terceiros

| Serviço | O que recebe | Política |
|---|---|---|
| Google Maps / Places (Google LLC) | Localização aproximada nas buscas; telemetria própria do SDK de mapas | [policies.google.com/privacy](https://policies.google.com/privacy) |
| Firebase (Google LLC) | Eventos de uso, diagnóstico/crash e identificador pseudônimo do app | [firebase.google.com/support/privacy](https://firebase.google.com/support/privacy) |
| Apple (App Store / StoreKit) | Processamento das compras de apoio | [apple.com/br/privacy](https://www.apple.com/br/privacy/) |
| Apple (MapKit/MKDirections) | Coordenadas de origem/destino no cálculo de rotas | [apple.com/br/privacy](https://www.apple.com/br/privacy/) |

## 3. Dados públicos exibidos pelo app

O app reproduz bases **públicas e oficiais** da ANP (fiscalização, cadastro de revendedores e preços), dos **Procons estaduais** e do sistema nacional **SINDEC/Senacon** (empresas autuadas e reclamações fundamentadas), do **IPEM-SP** (bombas certificadas) e conteúdo do Google (avaliações, notas), sempre com indicação de fonte e data.

- Dados de **empresas** (CNPJ, razão social) são exibidos como constam nas bases públicas.
- **CPFs** de revendedores pessoa física, quando presentes nas bases, são **mascarados** pelo app e pelo pipeline de dados — o documento completo não é redistribuído.
- As bases de defesa do consumidor contêm empresas (pessoas jurídicas), coletadas de portais de transparência e dados abertos (Lei de Acesso à Informação).
- As avaliações exibidas são de autoria de usuários do Google e permanecem sob responsabilidade de seus autores e da plataforma de origem.

## 4. Seus direitos (art. 18 da LGPD)

Você pode solicitar confirmação de tratamento, acesso, correção ou eliminação de dados pelo nosso canal de contato: **[fontinelles.com/contact](https://fontinelles.com/contact)**. Na prática: **desinstalar o app elimina os dados criados por você no aparelho**; a permissão de localização é revogável nos Ajustes do iOS; e você pode desativar a personalização de anúncios/rastreio do iOS a qualquer momento.

Se você é **titular de dados presentes nas bases públicas** exibidas (ex.: revendedor pessoa física), pode nos contatar para revisão da exibição — e, quanto às bases de origem, exercer seus direitos junto à ANP, ao Procon ou ao IPEM.

## 5. Crianças e adolescentes

O app é destinado a condutores habilitados e não é direcionado a menores de idade.

## 6. Alterações

Esta política pode ser atualizada; a versão vigente estará sempre neste endereço, com a data no topo.
