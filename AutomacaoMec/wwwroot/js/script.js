function preencherCampos() {
    for (let i = 0; i < textos.length; i++) {
        let campo = document.querySelector(`#campo_${i}`);
        if (campo) {
            campo.value = textos[i];
        }
    }
}
