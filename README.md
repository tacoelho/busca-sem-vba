
# Busca com sugestão automática sem VBA \o/

![Nerdy Dog](https://dogsaholic.com/wp-content/uploads/2018/08/nerdy-dog-with-a-laptop-810x515.png)

### Passo a passo

#### Dados


------------

Inserir quantas colunas de dados forem necessárias. Neste exemplo usaremos duas colunas com os dados **Cachorro** e **Raça**

![img1](https://i.imgur.com/4RmfHAb.jpg)

#### Preparando as colunas de consulta

------------
Agora precisamos incluir 3 colunas. Apenas por convenção e organização, podemos nomeá-las como "consulta1", "consulta2" e "consulta3".

![img2](https://i.imgur.com/7Si3MSM.jpg)


Na coluna "consulta1", vamos inserir a função **LINS()** para agilizar a numeração da quantidade de itens, caso contrário, teríamos que digitar individualmente o número da linha correspondente à localização do dado na coluna "Cachorro". Em uma planilha com milhares de linhas, isso seria inviável. Para saber mais sobre a função, consulte [aqui.](https://support.office.com/pt-br/article/lins-fun%C3%A7%C3%A3o-lins-b592593e-3fc2-47f2-bec1-bda493811597)

Note que a contagem começa a partir da 1ª célula com dado, não do cabeçalho, que não nos interessa:

![img3](https://i.imgur.com/rBR5iXA.jpg)

A fórmula é **=LINS($A$2:A2)** que nos retorna a posição numérica das células do intervalo desejado.

Agora iremos para a "Coluna2" e, neste momento, antes de inserirmos qualquer função, devemos inserir uma combo box na planilha (que será onde digitaremos o que desejamos buscar) e configurar algumas de suas propriedades. **Este é o passo que exige maior atenção, portanto siga fielmente o que vem a seguir:**

#### Configurando a Caixa de Combinação (Combo Box)

---

Siga até a aba Dados > Inserir > Controles ActiveX e selecione o controle Caixa de Combinação. No **Modo de Design** clique em **Propriedades** e faça os ajustes a seguir:
- Em **LinkedCell** digite a referência à célula onde os dados digitados na Caixa de Combinação irão ser replicados. Por exemplo, ao escolher a célula J5, tudo o que for digitado na Caixa de Combinação aparecerá na referida célula:

![img4](https://i.imgur.com/TW9rrmx.jpg)

- Em **ListFillRange** insira o intervalo dos dados. Neste exemplo, devemos preencher A1:B6.
- Em **MatchEntry** selecionar a opção **2 - fmMatchEntryNone** para evitar que a Caixa de Combinação tente adivinhar o que queremos buscar.

#### Retornando às colunas de consulta 

---

Na coluna "Consulta2" vamos inserir a seguinte fórmula: **=SE(ÉNÚM(LOCALIZAR($J$5;A2));C2;"")**, pois queremos encontrar a correspondência do que é digitado na Caixa de Combinação e exibido em J5 com os dados que estão na coluna A e verificar se encontram correspondência na coluna C. Como a referência é numérica, precisamos da função **=ÉNÚM**.

![img5](https://i.imgur.com/PKA9VYl.jpg)

Por fim, na "Coluna3", usaremos uma fórmula para retornar do menor para o maior uma correspondência da coluna C na coluna D e exibirmos somente o resultado encontrado. Ao digitarmos "Laika", por exemplo, a fórmula nos retorna a posição numérica de Laika na coluna D (Consulta2) que também corresponde à posição na coluna C (Consulta1). Portanto, temos um *match*. Caso contrário, o valor da célular seria vazio, por isso o uso de aspas em **=SEERRO**. A fórmula é **=SEERRO(MENOR($D$2:$D$6;C2);"")**.

![img6](https://i.imgur.com/EndkvpN.jpg)

#### Finalizando

---

Agora que já temos montada nossa base de dados, colunas adicionais para auxílio na consulta e as fórmulas que nos retornam a posição do dado procurado, devemos finalizar a planilha replicando nossa base de dados e inserindo uma última fórmula da seguinte forma:

1. Ocultar as colunas de dados e as de consulta. A planilha ficará com essa aparência:
2. Inserir a fórmula **=SEERRO(ÍNDICE($A$2:$A$6;$E2;COLS($J$8:J8));"")** onde quiser que uma cópia das colunas "Cachorro" e "Raça" apareçam. Esta fórmula irá replicar a coluna "Cachorro" ($A$2:$A$6) em $J$8:J8. Para que replicar a coluna "Raça" substitua $A$2:$A$6 por $B$2:$B$6 e $J$8:J8 por $K$8:K8. A planilha ficará com a seguinte aparência:

![img7](https://i.imgur.com/8dTAlBv.jpg)

Basicamente, o que esta fórmula faz, é buscar no intervalo informado aquele dado que filtramos nas colunas de consulta até chegar à posição exata exibida na coluna "Consulta3" ($E2) e exibir o resultado do que digitamos na Caixa de Combinação nas colunas replicadas ($J$8:J8 e $K$8:K8). Para que o resultado da busca não fique aparecendo isoladamente, oculte a linha, porém, nesse caso, ponha a célula alvo fora do intervalo de linhas dos dados principais para não ocultar uma linha desses dados também.

Reexibindo as colunas, podemos ver melhor o funcionamento e integração das fórmulas:

![img8](https://i.imgur.com/MPJE13P.jpg)

# Então é isso, pessoal! Uma ferramenta útil pra você impressionar o chefe e sem digitar uma linha de código VBA :^]

