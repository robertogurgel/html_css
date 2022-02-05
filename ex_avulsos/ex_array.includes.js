//array.includes(item)
// verifica se o item tem na array e rotorna true ou false

let frutas=['uva','banana','limão'];
let compras=['uva','pera','banana','maça','limão'];

for (let item of compras){

    console.log(item  + ': ' + frutas.includes(item))
    
}