var Lettera = function(infoLettera, dataInvio, pratiche) {
        this.infoLettera = infoLettera,
        this.dataInvio = dataInvio,
        this.pratiche = pratiche
        this.numeroPratiche = function(){            
            return pratiche.length
        }
        this.print = function(){        
        Logger.log('Stampo la lettera per le seguenti ' + this.numeroPratiche() + ' pratiche: ' + JSON.stringify(this.pratiche))
        }
}

function creaLettera(){

    var infoLettereArray2D = sheetInfoLettere.getDataRange().getValues()
    var infoLettere = ObjApp.rangeToObjectsNoCamel(infoLettereArray2D)
    Logger.log(infoLettere)
    var dataInvio = '01/06/2017'

    var praticheArray2D = ssFileEsportazione.getDataRange().getValues()


    var praticheArrayObj = ObjApp.rangeToObjectsNoCamel(praticheArray2D)
    var infoLettera = infoLettere[0]
    var pratiche = praticheArrayObj.filter(function (el) {
        Logger.log(el['Account'])
        return el['AccountsStatus'] == infoLettera.Status
    });
    var lettera = new Lettera(infoLettere[0], dataInvio, pratiche);
    lettera.print()
}
