### Aviso: Este código é uma adaptação gerada com o auxílio de inteligência artificial, e seu criador não possui conhecimento de programação.
### Código original: https://github.com/drxu-drxu-creative-sinner/alocacaocpnu/tree/main

### **Relatório de Metodologia de Alocação de Candidatos**

**Assunto:** Análise funcional e algorítmica do script alocacao.py para alocação de candidatos a vagas.

#### **1\. Resumo Executivo**

O script alocacao.py implementa um sistema de alocação de candidatos a vagas em diversos órgãos, com base em múltiplas regras de negócio e preferências. O objetivo é preencher um quadro de vagas pré-definido (VAGAS\_POR\_CARGO), respeitando as notas dos candidatos, as cotas (PCD, PPP, PI), e a ordem de preferência explícita de cada candidato.  
O processo é executado em duas fases principais:

1. **Fase 1 (Classificação Preliminar):** Uma classificação isolada por cargo é realizada para identificar todos os candidatos elegíveis dentro do número de vagas, com base puramente no mérito (nota e critérios de desempate).  
2. **Fase 2 (Resolução de Conflitos):** Um processo iterativo é executado para resolver conflitos de candidatos aprovados em múltiplas vagas. Este processo utiliza a ordem de preferência do candidato (Ordem\_Pref) como o principal critério de decisão.

O paradigma matemático subjacente é uma variação do **Algoritmo de Emparelhamento Estável** (Stable Matching Algorithm), comumente aplicado a problemas de alocação (ex: "College Admissions Problem" ou "Hospital-Resident Problem").

#### **2\. Fase 1: Preparação e Classificação Preliminar**

O script inicia com a preparação dos dados e uma primeira alocação baseada exclusivamente no mérito dentro de cada cargo.  
2.1. Ingestão e Validação de Dados  
O script lê uma planilha Excel (b7p.xlsx), define os caminhos de saída e carrega a estrutura de vagas (VAGAS\_POR\_CARGO). Este dicionário é a "fonte da verdade" que define o número exato de vagas (AMPLA, PPP, PCD, PI) para cada combinação de Órgão/Cargo.  
Os dados dos candidatos são higienizados e normalizados:

* Colunas de cotas são consolidadas em um único campo (Cota).  
* Campos numéricos (Nota Final, Ordem\_Pref, Posição Real) são convertidos para o tipo numérico adequado.  
* É criado um universo de candidatos válidos, filtrando apenas aqueles que manifestaram interesse (Interesse\_norm \== 'sim') e possuem Nota Final e Ordem\_Pref válidas.

2.2. Métrica de Classificação (Preferência da Vaga)  
Para estabelecer a ordem de mérito, o script cria uma "tupla de ordenamento" (rank\_tuple) para cada candidato em cada cargo. Esta tupla define a lista de preferência da vaga (ou seja, a ordem de classificação):

1. **Nota Final** (negativa, para ordenar da maior para a menor).  
2. **Posição Real** (critério de desempate do edital).  
3. **CandID** (critério final de desempate para garantir ordenamento único).

Notavelmente, a Ordem\_Pref do candidato **não** é usada nesta etapa de classificação.  
2.3. Alocação Preliminar por Cargo  
O script itera sobre cada cargo individualmente:

1. **Preenchimento de Cotas (Reservas):** O sistema aloca primeiro os candidatos elegíveis para as vagas reservadas (PI, PCD, PPP), seguindo estritamente a ordem do rank\_tuple.  
2. **Preenchimento da Ampla Concorrência:** As vagas restantes (AMPLA, mais quaisquer vagas de cota não preenchidas) são então preenchidas pelos candidatos com melhor rank\_tuple, independentemente da cota do candidato.  
3. **Criação de Listas de Espera (backlog):** Os candidatos que se classificaram para o cargo mas ficaram fora do número de vagas são armazenados em uma lista de espera (backlog) específica para aquele cargo.

Ao final desta fase, um candidato pode estar classificado (aprovado) em múltiplas vagas.

#### **3\. Fase 2: Metodologia de Alocação e Resolução de Conflitos**

Esta fase resolve os conflitos da Fase 1, garantindo que cada candidato seja alocado em apenas uma vaga, respeitando sua preferência.  
3.1. Paradigma: Algoritmo de Emparelhamento Estável  
O núcleo do processo é um algoritmo de emparelhamento estável. Este problema matemático envolve dois conjuntos de "atores" com listas de preferências:

* **Candidatos:** Sua preferência é a Ordem\_Pref (ex: 1º, 2º, 3º lugar).  
* **Vagas:** Sua preferência é o rank\_tuple (ex: maior nota, segunda maior nota, etc.).

O objetivo do script é encontrar uma "alocação estável" que prioriza a escolha do candidato.  
3.2. Processo Iterativo de Estabilização  
O script entra em um loop (com um limite de 8 iterações, interrompido se não houver mudanças) que alterna entre duas funções: resolve\_global e backfill\_all.

1. **Resolução de Conflitos (resolve\_global):**  
   * O sistema identifica todos os candidatos que receberam múltiplas "ofertas" (ou seja, foram classificados em mais de uma vaga na Fase 1).  
   * Para cada um desses candidatos, o script consulta sua Ordem\_Pref.  
   * O candidato é **mantido** apenas na vaga correspondente à sua **menor Ordem\_Pref** (sua maior preferência).  
   * O candidato é **removido** de todas as outras vagas onde havia sido classificado.  
2. **Preenchimento de Vagas (backfill\_all):**  
   * Quando um candidato é removido de uma vaga (pelo resolve\_global), essa vaga fica ociosa.  
   * A função backfill\_all é acionada para "puxar" o próximo candidato da lista de espera (backlog) daquele cargo, respeitando as regras de cota da vaga que ficou livre.  
   * Este candidato "puxado" pode, por sua vez, já estar alocado em outra vaga (de preferência menor), criando um novo conflito a ser resolvido na próxima iteração.

Este ciclo de "resolver conflitos" e "preencher vagas" continua até que nenhum candidato mude de vaga, resultando em um sistema estável.

#### **4\. Fase 3: Geração de Relatórios e Auditoria**

Após a estabilização, o script gera os seguintes artefatos de saída:

* **detalhes\_alocacao...csv:** A lista final e detalhada de todos os candidatos alocados, contendo o órgão, cargo, dados do candidato e o tipo de vaga (cota) utilizada.  
* **resultado\_alocacao...xlsx:** Um arquivo Excel com múltiplas abas:  
  * Detalhes...: Cópia do CSV de detalhes.  
  * Resumo Vagas: Um comparativo entre o número de vagas planejado (do VAGAS\_POR\_CARGO) e o número de vagas efetivamente preenchidas por tipo de cota.  
  * Contagem por Cota: Um resumo quantitativo de alocados por cargo e tipo de vaga.  
* **nao\_alocados...csv:** Uma lista de candidatos que, embora pertencessem ao universo válido, não foram classificados em nenhuma vaga ao final do processo.  
* **auditoria\_desempates...csv:** Um relatório de verificação que confirma se, em casos de empate na Nota Final dentro de um mesmo cargo/tipo de vaga, a ordenação final obedeceu ao critério de desempate Posição Real.

#### **5\. Conclusão**

O script alocacao.py implementa um método robusto e justo de alocação. Ele garante que a classificação de mérito (rank\_tuple) seja usada para preencher as vagas, enquanto assegura que, em casos de múltiplas aprovações, a preferência explícita do candidato (Ordem\_Pref) seja o critério soberano de decisão. O resultado é um emparelhamento estável e auditável.
