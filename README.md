# Simulador Educacional de Investimentos em FIIs

Projeto desenvolvido como parte de um desafio da **DIO**, aplicando conceitos de Excel na criação de uma ferramenta prática para simulação de investimentos em Fundos de Investimento Imobiliário (FIIs).

O simulador permite informar salário, forma de aporte, prazo e taxas mensais para estimar o patrimônio acumulado, os dividendos mensais e a distribuição do aporte conforme o perfil do investidor.

> **Aviso:** este projeto possui finalidade exclusivamente educacional. Os resultados são estimativas matemáticas e não representam promessa de rentabilidade, recomendação financeira ou previsão de mercado.

## Objetivos do projeto

- Criar uma ferramenta de simulação de investimentos em Excel;
- Aplicar cálculos financeiros de capitalização mensal;
- Estimar patrimônio, dividendos e rendimento acumulado;
- Comparar diferentes prazos de investimento;
- Distribuir o aporte conforme o perfil do investidor;
- Criar uma interface simples e segura;
- Documentar os conhecimentos adquiridos.

## Funcionalidades

- Cálculo do aporte por valor mensal ou percentual do salário;
- Projeção do patrimônio com capitalização mensal;
- Estimativa dos dividendos mensais;
- Cálculo do total aportado;
- Cálculo do rendimento acumulado;
- Comparação de cenários para 2, 5, 10, 20 e 30 anos;
- Seleção dos perfis Conservador, Moderado e Agressivo;
- Distribuição automática entre diferentes tipos de FIIs;
- Gráficos de evolução do patrimônio e dos dividendos;
- Listas suspensas para padronizar as escolhas;
- Proteção dos textos e das fórmulas.

## Como utilizar

1. Abra a planilha no Microsoft Excel.
2. Acesse a aba **Simulador**.
3. Preencha somente as células amarelas da área **Configurações**:
   - salário mensal;
   - forma do aporte;
   - percentual do salário;
   - valor mensal informado;
   - prazo do investimento;
   - taxa de rendimento mensal;
   - taxa mensal estimada de dividendos;
   - perfil do investidor.
4. Consulte os resultados calculados automaticamente.
5. Compare os diferentes cenários de prazo.
6. Analise os gráficos de patrimônio e dividendos.
7. Verifique a distribuição sugerida para o perfil selecionado.

## Estrutura da planilha

### Aba Simulador

É a interface principal do projeto e contém:

- configurações informadas pelo usuário;
- resultados da simulação;
- projeções para diferentes prazos;
- distribuição mensal por tipo de FII;
- gráficos de patrimônio e dividendos.

### Aba Perfis

É a base auxiliar do simulador e contém:

- perfis de investidor;
- tipos de FIIs;
- percentuais de distribuição;
- chaves utilizadas nas pesquisas;
- controles para verificar se cada perfil totaliza 100%.

Essa aba está protegida para evitar alterações que possam comprometer os cálculos.

## Perfis e distribuição

| Tipo de FII | Conservador | Moderado | Agressivo |
| --- | ---: | ---: | ---: |
| Papel | 30% | 32% | 50% |
| Tijolo | 50% | 35% | 10% |
| Híbridos | 10% | 8% | 5% |
| FOFs | 10% | 5% | 5% |
| Desenvolvimento | 0% | 10% | 20% |
| Hotelarias | 0% | 10% | 10% |
| **Total** | **100%** | **100%** | **100%** |

Os percentuais são premissas educacionais utilizadas para demonstrar a distribuição automática dos aportes. Eles não constituem recomendação de carteira.

## Principais cálculos

### Sugestão de aporte pelo salário

```excel
=salario*percentual_salario
```

Calcula quanto seria investido quando o usuário escolhe definir o aporte como percentual do salário.

### Aporte mensal efetivo

```excel
=IF(forma_aporte="% do salário",salario*percentual_salario,valor_mensal)
```

Escolhe automaticamente entre o percentual do salário e o valor mensal informado.### Patrimônio acumulado

```excel
=SE(qtd_anos=0;0;VF(taxa_mensal;qtd_anos*12;-aporte))```

Utiliza a função de valor futuro para calcular o patrimônio com aportes e capitalização mensais.

### Dividendos mensais estimados

```excel
=patrimonio*rendimento_carteira
```

Aplica a taxa mensal estimada de dividendos ao patrimônio projetado.

### Total aportado

```excel
=aporte*qtd_anos*12
```

Multiplica o aporte mensal pela quantidade total de meses.

### Rendimento acumulado

```excel
=patrimonio-total_aportado
```

Calcula a diferença entre o patrimônio projetado e a soma dos aportes.

### Distribuição por tipo de FII

```excel
=aporte*percentual_do_perfil
```

Distribui o aporte mensal conforme os percentuais do perfil selecionado.

> Dependendo do idioma do Excel, algumas funções podem aparecer traduzidas, como `IF`/`SE` e `FV`/`VF`.

## Tecnologias e conceitos aplicados

- Microsoft Excel;
- fórmulas financeiras;
- juros compostos;
- capitalização mensal;
- referências nomeadas;
- funções condicionais;
- funções de pesquisa;
- validação de dados;
- listas suspensas;
- formatação de moedas e percentuais;
- gráficos de colunas;
- proteção de células e fórmulas;
- separação entre interface e base auxiliar;
- controle de consistência dos percentuais.

## Proteção da planilha

Na aba **Simulador**, somente os campos de entrada permanecem desbloqueados. Textos, fórmulas, resultados e gráficos estão protegidos contra alterações acidentais.

Para realizar manutenção:

1. Acesse **Revisão > Desproteger Planilha**.
2. Faça as alterações necessárias.
3. Acesse **Revisão > Proteger Planilha**.
4. Permita apenas a seleção de células desbloqueadas.

A proteção atual não utiliza senha.

## Aprendizados

O desenvolvimento deste projeto permitiu praticar:

- transformação de requisitos em uma ferramenta funcional;
- organização de entradas, premissas e resultados;
- aplicação de juros compostos;
- utilização de fórmulas financeiras;
- criação de cálculos que respondem às escolhas do usuário;
- validação de dados para reduzir erros;
- separação dos dados auxiliares da interface principal;
- controle para garantir que as distribuições totalizem 100%;
- proteção do modelo sem impedir o preenchimento;
- apresentação dos resultados por meio de tabelas e gráficos;
- documentação técnica para publicação no GitHub.

## Melhorias futuras

- Criar cenários otimista, base e pessimista;
- Permitir taxas diferentes para cada tipo de FII;
- Considerar a inflação;
- Permitir crescimento anual do aporte;
- Adicionar reinvestimento configurável dos dividendos;
- Considerar impostos, custos e vacância;
- Comparar os FIIs com outras classes de ativos;
- Importar indicadores de fontes externas;
- Criar um painel com indicadores adicionais;
- Adicionar testes automáticos de consistência.

## Limitações

- As taxas permanecem constantes durante cada simulação;
- Impostos, custos, vacância e oscilações de mercado não são considerados;
- Os percentuais dos perfis são premissas didáticas;
- Os resultados dependem dos valores informados pelo usuário;
- Resultados projetados não garantem desempenho futuro.

## Autor

Desenvolvido por **Edson Deveza** como parte da formação na **DIO**.

## Licença

Este projeto está disponível sob a licença MIT.

Copyright (c) 2026 Edson Deveza

É concedida permissão, gratuitamente, para usar, copiar, modificar e distribuir este projeto, desde que o aviso de direitos autorais seja mantido.

O projeto é fornecido “no estado em que se encontra”, sem garantias de qualquer tipo.
