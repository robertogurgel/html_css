function carregar() { 
    var msg = window.document.getElementById('msg')
    var img = window.document.getElementById('imagem')
    var data = new Date()
    var hora = data.getHours() 
    
    //para testar sem pegar a data automatica do sistema
    //var hora= '22'
    msg.innerHTML = `Agora são ${hora} horas.`

    if (hora>=0 && hora<12) {
        img.src = 'cajoni_310x310.png'

        document.body.style.background = '#fff200'
    } else if (hora>=12 && hora<=18) {
        img.src = 'cajoni_310x310.png'
        document.body.style.background = '#ed1c24'
    } else  {
        img.src = 'cajoni_310x310.png'
        document.body.style.background = '#00a651'
    }

} 
    


