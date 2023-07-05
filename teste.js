function pesquisarNoBing() {
  var documento = DocumentApp.getActiveDocument();
  var termoPesquisa = documento.getBody().getText();
  
  var url = 'https://api.cognitive.microsoft.com/bing/v7.0/search?q=' + encodeURIComponent(termoPesquisa);
  var chaveBing = ''; //chave da API do Bing

  var respostaBing = UrlFetchApp.fetch(url, {
    headers: {
      'Ocp-Apim-Subscription-Key': chaveBing
    }
  });
  
  var resultadosBing = JSON.parse(respostaBing.getContentText());
  
  // Faça o processamento dos resultados da pesquisa do Bing aqui
  var itensBing = resultadosBing.webPages.value.slice(0, 5);
  
  var tabelaDados = [];
  
  // Monta a tabela com os cabeçalhos
  tabelaDados.push(["Código", "ID", "Nome", "Estoque Atual", "Preço"]);
  
  // Adiciona os dados dos produtos na tabela
  itensBing.forEach(function(itemBing) {
    var codigo = ""; // Código do produto
    var id = ""; // ID do produto
    var nome = ""; // Nome do produto
    var estoqueAtual = ""; // Estoque atual do produto
    var preco = ""; // Preço do produto
    
    // Faz a busca dos dados do produto na sua fonte de dados (por exemplo, banco de dados, API, etc.)
    // Utilize as informações do itemBing (título, URL, etc.) para encontrar os dados corretos do produto
    
    // Adiciona os dados na tabela
    tabelaDados.push([codigo, id, nome, estoqueAtual, preco]);
  });
  
  // Insere a tabela no documento do Google Docs
  documento.getBody().appendTable(tabelaDados);
}
