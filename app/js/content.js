((window, browser) => {
    const hostName = "http://127.0.0.1:51003/";
    console.info("Add BarsPy object:");
    console.info("Вывод на принтер: ");
    console.info('   BarsPy.send({"GetPrinterList":1},function(dat){console.log(dat);}) // получить список принтеров установленных в системе');
    console.info('   BarsPy.send({"Print":"<h1>Привет Мир-HelloWorld</h1>","widthPage":300,"heightPage":100,"PrinterName":"Microsoft XPS Document Writer"},function(dat){console.log(dat);})');
    console.info('   BarsPy.send({"Print":"<h1>Привет Мир-HelloWorld</h1>"}) // отправека на печать без получения ответа ');



    function includeScript(path) {
        return new Promise((resolve, reject) => {
            var req = new XMLHttpRequest();
            req.onreadystatechange = function () {
                if (req.readyState == 4) {
                    if (req.status >= 200 && req.status < 300) {
                        resolve(req.responseText);
                    } else {
                        reject(req);
                    }
                }
            };
            req.open("GET", browser.extension.getURL(path));
            req.send();
        });
    }

    function injectScript(text) {
        var s = document.createElement('script');
        s.appendChild(document.createTextNode(text));
        s.onload = function () {
            this.parentNode.removeChild(this);
        };
        (document.head || document.documentElement).appendChild(s);
    }


    includeScript("js/adapter.js")
        .then((script) => {
            injectScript("" +
                "(function() {" +
                "  const host = '" + hostName + "';" +
                script +
                "  window['BarsPy'] = new BarsPyAdapter(host);" +
                "})()");
        })
        .catch((error) => {
            console.warn("Could not initialize \"BARS Browser Python Adapter\" extension", error);
        })

})(window, window.chrome ? chrome : browser);