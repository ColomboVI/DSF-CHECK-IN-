/**************************************************************************\
 * Copyright (C) 2018 by Synergic Partners                                 *
 *                                                                         *
 * author     : Borja Durán                                                *
 * description:                                                            *
 * - clase para hacer resgistros de acciones dentro del excel en la pestaña LOG
 *                                                                         *
 * TODO                                                                    *
 * ====                                                                    *
 * - .....                                                                 *
 * ----------------------------------------------------------------------- *
 * This program is not free software; you can not : (a) copy or use the    *
 * Software in any manner except as expressly permitted by SynergicPartners*
 * (b) transfer, sell, rent, lease, lend, distribute, or sublicense the    *
 * Software to any third party; (c)  reverse engineer, disassemble, or     *
 * decompile the Software; (d) alter, modify, enhance or prepare any       *
 * derivative work from or of the Software; (e) redistribute it and/or     *
 * modify it without prior, written approval from Synergic Partners.       *
\***************************************************************************/
function Logging(disparador) {

  if (typeof Logging.instancia === 'object') {
        return Logging.instancia;
    }

  this._disparador='# '+disparador;
  this._ss=SpreadsheetApp.getActive().getSheetByName('LOG');


    this.newEventTexts=function(traza, resultado)
    {
        this._ss.getRange(this._ss.getLastRow()+1,1,1,3).setValues([[(new Date()).toLocaleString(),this._disparador+': '+traza,resultado]]);
    }
    this.newEventTextFormula=function(traza, formula)
    {
        var i=this._ss.getLastRow()+1;
        this._ss.getRange(i,1,1,2).setValues([[(new Date()).toLocaleString(),this._disparador+': '+traza]]);
        this._ss.getRange(i,3,1,1).setFormula(formula);
    }

    Logging.instancia = this;
}

function pruebaLogging()
{
var logger=new Logging('MANUAL (pruebaLogging)');
logger.newEventTexts('texto1','texto2');
logger.newEventTextFormula('texto1','=HYPERLINK("http://www.google.com/";"Google")');
}
