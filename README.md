
# Busca com sugestão automática sem VBA \o/

![Nerdy Dog](https://dogsaholic.com/wp-content/uploads/2018/08/nerdy-dog-with-a-laptop-810x515.png)

### Passo a passo

###### Dados


------------

<p>Inserir quantas colunas de dados forem necessárias. Neste exemplo usaremos duas colunas com os dados **Cachorro** e **Raça**</p>

![img1](https://i.imgur.com/4RmfHAb.jpg)

###### Preparando as colunas de consulta

------------
Agora precisamos incluir 3 colunas. Apenas por convenção e organização, podemos nomeá-las como "consulta1", "consulta2" e "consulta3".

![img2](https://i.imgur.com/7Si3MSM.jpg)


Na coluna "consulta1", vamos inserir a função **LINS()** para agilizar a numeração da quantidade de itens, caso contrário, teríamos que digitar individualmente o número da linha correspondente à localização do dado na coluna "Cachorro". Em uma planilha com milhares de linhas, isso seria inviável. Para saber mais sobre a função dados, consulte [aqui.](https://support.office.com/pt-br/article/lins-fun%C3%A7%C3%A3o-lins-b592593e-3fc2-47f2-bec1-bda493811597)

Note que a contagem começa a partir da 1ª célula com dado, não do cabeçalho, que não nos interessa:

![img3](https://i.imgur.com/rBR5iXA.jpg)

A fórmula é **=LINS($A$2:A2)** que nos retorna a posição numérica das células do intervalo desejado.

Agora iremos para a "Coluna2" e, neste momento, antes de inserirmos qualquer função, devemos inserir uma combo box na planilha (que será onde digitaremos o que desejamos buscar) e configurar algumas de suas propriedades. **Este é o passo que exige maior atenção, portanto siga fielmente o que vem a seguir: **
