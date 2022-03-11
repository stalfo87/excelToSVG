
const convert = require('xml-js');

const xml = `<cjuego><juego>Bingo Story</juego><solicitudJuegosOS><tipoMaquina>BLUE</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego>
<cjuego><juego>La Calaca Bingo</juego><solicitudJuegosOS><tipoMaquina>WBR2</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>Fairy Fantasy</juego><solicitudJuegosOS><tipoMaquina>SLEIC</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>DISCOBALL</juego><solicitudJuegosOS><tipoMaquina>BLUE</tipoMaquina><cantidad>2</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>Fire Bingo</juego><solicitudJuegosOS><tipoMaquina>WBR1</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>DISCOBALL</juego><solicitudJuegosOS><tipoMaquina>BLUE</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>Fire Bingo</juego><solicitudJuegosOS><tipoMaquina>WBR1</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>Bingo 
Story</juego><solicitudJuegosOS><tipoMaquina>BLUE</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego>`

console.log(convert.xml2js(null, {compact:true, ignoreDeclaration:true}))