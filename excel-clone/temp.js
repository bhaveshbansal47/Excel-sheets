var arr = [];
for(var i = 0; i < 100; i++){
    
    arr.push(Math.floor(Math.random()
    *6 + 1));
}
var six = [];
for(var i = 0; i < 6; i++){
    six.push(arr[Math.floor(Math.random()*100)]);
}

console.log(six[Math.floor(Math.random()*6)]);