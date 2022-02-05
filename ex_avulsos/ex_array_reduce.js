//array.reduce (soma os itens de uma array)
let hours=[8,8,6,0,8,8,12];

console.log(hours);
let totalHours = hours.reduce((sum,today)=>{return sum + today},0);
console.log('Total Hours: ' + totalHours)

let cost=[4,8,15,16,23,42];
console.log(cost)
let totalCost = cost.reduce((sum,price)=>{return sum + price},0);
console.log('Total Preços :' + totalCost);