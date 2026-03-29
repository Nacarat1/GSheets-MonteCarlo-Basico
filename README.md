# GSheets-MonteCarlo-Basico
Este é um código em AppScript que irá permitir que rode Simulações Monte Carlo (Distribuição PERT) direto no seu google sheets, facilitando a entrada em estimativas e simulações de cenários. 

Ideal para gerentes de projeto, engenheiros, analistas financeiros e qualquer profissional que precise calcular cenários de risco, prazos e orçamentos fugindo da velha estimativa de "ponto único".


## Funcionalidades

* **Modelagem Híbrida:** Utiliza o histórico de projetos anteriores (se disponível) extraindo a mediana, mínimo e máximo automaticamente.
* **Estimativa de 3 Pontos (PERT):** Se não houver histórico, utiliza sua estimativa de Cenário Otimista, Mais Provável e Pessimista, aplicando uma Curva Beta contínua.
* **1.000 Iterações Automáticas:** Roda mil cenários possíveis em segundos.
* **Gráficos Executivos:** Gera automaticamente uma **Curva S** (Probabilidade Acumulada) e um **Histograma de Densidade** para fácil visualização do risco.
* **Análise de Alvo:** Permite que você insira um "Chute/Orçamento Alvo" e retorna a probabilidade exata daquele cenário acontecer.
* **Interface Pronta:** O próprio script formata a planilha com larguras, cores e títulos para facilitar a usabilidade.

## Como Instalar no seu Google Sheets

Você não precisa instalar nenhum software. Basta seguir os passos abaixo:

1. Abra uma nova planilha em branco no [Google Sheets](https://sheets.google.com).
2. No menu superior, clique em **Extensões** > **Apps Script**.
3. Apague qualquer código que estiver no editor e **cole todo o código** do arquivo `MonteCarlo.js` deste repositório.
4. Clique no ícone de disquete (Salvar) ou pressione `Ctrl + S`.
5. Feche a aba do Apps Script e volte para a sua planilha.
6. **Atualize a página (F5)**. Você verá um novo menu chamado **"🎲 Simulação Monte Carlo"** ao lado de "Ajuda".

*(Nota: Na primeira vez que você rodar, o Google pedirá uma autorização de segurança. Basta clicar em "Continuar", escolher sua conta Google, ir em "Avançado" e clicar em "Acessar script").*

##  Como Usar

1. No menu superior, clique em **🎲 Simulação Monte Carlo** > **1. Configurar Planilha Padrão**. O script vai desenhar o template na sua tela.
2. **Preencha os dados (Células Verdes):**
   * **Caso possua um histórico:** Cole seus dados reais na Coluna A (`Histórico`), Deixe a estimativa vazia.
   * **Caso possua uma estimativa de 3 pontos:** Deixe a Coluna A vazia e preencha suas estimativas de 3 pontos nas células `C2`, `C3` e `C4`.
3. Defina sua unidade de medida na célula `F1` (ex: R$, Dias, %, kg).
4. Insira o seu "Valor Alvo" (o orçamento ou prazo que você quer testar) na célula `F5`.
5. Vá novamente no menu **🎲 Simulação Monte Carlo** e clique em **2. Rodar Simulação e Gerar Gráficos**.

Pronto! Os percentis (P10, P50, P90), a probabilidade do seu alvo e os gráficos executivos aparecerão automaticamente na tela.

##  Tecnologias Utilizadas
* Google Apps Script (JavaScript)
* Google Sheets API
* Matemática Estatística (Distribuição Beta / PERT e Geração de Números Pseudoaleatórios)

##  Contribuições
Sinta-se à vontade para abrir *Issues* relatando bugs ou *Pull Requests* sugerindo melhorias (como novos percentis, distribuições diferentes como Normal/Lognormal, etc).
