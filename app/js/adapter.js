/*
 * Copyright (c) 2020, BARS Group. All rights reserved.
 */

/**
 * BarsPyAdapter
 * Позволяет отправлять\получать сообщения от расширения.
 *
 * @author Myasnikov Ivan Alekcandrovich  based  Vadim Shushlyakov
 */
class BarsPyAdapter {
    constructor(host) {
        this.host = host;
    }
    send(messageObject,FunCallBack) {
		var host = "http://127.0.0.1:51003/";
		var cspBindRequestCall = new XMLHttpRequest();
		cspBindRequestCall.open('GET',host, true);
		if (typeof FunCallBack === 'function'){ 
			 cspBindRequestCall.onreadystatechange = function() {
			 if (this.readyState == 4 && this.status == 200) {
				if (typeof FunCallBack === 'function'){
				try {
					FunCallBack(JSON.parse(decodeURIComponent(this.responseText)));
				} catch (err) {
					FunCallBack(decodeURIComponent(this.responseText));
				}
			  }
			 };
		   };
		}
		cspBindRequestCall.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		if (typeof messageObject == "object"){
			for (var prop in messageObject) {
			   cspBindRequestCall.setRequestHeader("X-My-"+prop, encodeURI(messageObject[prop]));
			}
		} else{
			cspBindRequestCall.setRequestHeader("X-My-message", encodeURI(messageObject));
		}
		cspBindRequestCall.send();
		return cspBindRequestCall; 
    }

    /**
     * Возвращает уникальный идентификатор
     * @returns {string}
     */
    _genId() {
        // https://stackoverflow.com/a/2117523
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'
            .replace(/[xy]/g, function (c) {
                let r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
                return v.toString(16);
            });
    }
}




