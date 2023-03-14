const fs = require('fs');
const rateLimit = require('axios-rate-limit');
const mammoth = require('mammoth');
const axios = require('axios');
const cheerio = require("cheerio");
const htmlToDocx = require('html-docx-js');
const apiKey = 'CHAVE_API';
const apiUrl = 'https://api.openai.com/v1/engines/text-davinci-003/completions';
const arquivo_a_ser_editado = 'C:\\Temp\\M.docx';//mude para o diretorio do arquivo que vai ser editado
const saida_arquivo_editado = 'C:\\Temp\\AS.docx';
var perguntaPadrão =  "'Corrija gramaticalmente o seguinte texto:";

var textArray = [];
var $;
const maximo_tentativas = 3;
const api = rateLimit(axios.create(), { maxRequests: 7, perMilliseconds: 1000 });//configura 7 segundos entre uma requisição e outra, essa é a taxa do openia em média

async function main(textoPorParagrafo) {

    try {
        const sucesso = [];

        for (let i = 0; i < textoPorParagrafo.length; i++) {

            var checkPoint = true; //essa variavel controla a entrada e saida do while

            const paramGpt = {
                prompt: perguntaPadrão + textoPorParagrafo[i] + "'",
                temperature: 0,
                max_tokens: parseInt(parseInt(perguntaPadrão.length + 10, 10) + parseInt(textoPorParagrafo[i].length + 20, 10), 10), //coloca o numero de tokens a quantidade de caracter mais 10 
                top_p: 1,
                frequency_penalty: 0,
                presence_penalty: 0.2,
                n: 1//textoPorParagrafo[i] == "texto no arquivo" ? 1000 : 1, //aqui  vc pode usar isso pra forçar um erro
                
            };

            let tentativas = 0;
            let continuar = true;
            /* esse while é para tentar algumas vezes caso der erro para enviar para IA gpt, se der erro ele salva somente o que deu certo */
            while (tentativas <= maximo_tentativas && checkPoint === true) {
                try {
                    const response = await api.post(apiUrl, paramGpt, {
                        headers: {
                            'Content-Type': 'application/json',
                            'Authorization': `Bearer ${apiKey}`
                        }
                    })
                    sucesso.push(response.data);
                    checkPoint = false; 
                } catch (error) {
                    console.log(`Erro na tentativa ${tentativas + 1} de ${i + 1}ª requisição:`, error.message);
                    tentativas++;
                    checkPoint = true;
                    if (tentativas === maximo_tentativas) {
                        continuar = false; // Interrompe o loop de tentativas se já foram feitas as tentivas na verivel 
                        console.log(`A requisição ${i + 1} seu erro após ${maximo_tentativas} tentativas. Economizei seu dinheiro, de nada huauauhua`);
                    }
                }
            }
            if (!continuar) {
                break;
            }
        }

        sucesso.forEach((completion, index) => {
            
            $('body *').each((i, element) => { //percorre o html para trocar os textos e manter os elementos intactos

                if (textoPorParagrafo[index] == $(element).text()) { //verifica se o texto é igual ao vetor inicial
                    const text = $(element).text()
                    if (text.includes(text)) {
                        const newHtml = $(element).html().replace(textoPorParagrafo[index], completion.choices[0].text)
                        $(element).html(newHtml)
                        return false; // sai do loop para não procurar nos outros elementos
                    }
                }
            })

        });


        const novoHtml = $.html();
        const converted = htmlToDocx.asBlob(novoHtml); // converte o HTML em Blob do arquivo DOCX

        const buffer = Buffer.from(await converted.arrayBuffer()); // converte o Blob em buffer
        fs.writeFileSync(saida_arquivo_editado, buffer); //salva no disco
        console.log("Arquivo salvo com sucesso")

        return sucesso;
    } catch (error) {
        return `Erro ao salvar o arquivo ${saida_arquivo_editado}`;
        

    }
}

mammoth.convertToHtml({ path: arquivo_a_ser_editado })
    .then((result) => {

        const html = result.value;
        $ = cheerio.load(html);

        $("p").each(function () {
            textArray.push($(this).text()); //transforma em paragrafo para enviar pra a ultima fronteira(IA)
        });

        main(textArray)
            .then((sucesso) => console.log('A função foi chamada :', sucesso))
            .catch((error) => console.log('Erro ao fazer as requisições:', error.message));

    })


