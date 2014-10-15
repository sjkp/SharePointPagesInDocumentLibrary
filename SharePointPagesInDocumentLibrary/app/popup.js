function onPageDetailsReceived(pageDetails) {
    document.getElementById('listTitle').value = pageDetails.title;
    document.getElementById('webUrl').value = pageDetails.url;    
}


var wopiframe = "/_layouts/15/WopiFrame.aspx?sourcedoc={0}&action=interactivepreview&wdSmallView=1"

var getFiles = "/_api/web/lists/GetByTitle('{0}')/items?$expand=File"
//var owa = "https://euc-word-view.officeapps.live.com/wv/docdatahandler.ashx?WOPIsrc=https%3A%2F%2Fsjkpdevs%2Dmy%2Esharepoint%2Ecom%2Fpersonal%2Fadmin%5Fsjkpdevs%5Fonmicrosoft%5Fcom%2F%5Fvti%5Fbin%2Fwopi%2Eashx%2Ffiles%2F0e8cd99eb770461c9d1d03e4a55bb8ac&access_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjdrNTBKdl9vYVloQ3lHcDBOVUlQdkFYQ3l0TSJ9%2EeyJhdWQiOiJ3b3BpL3Nqa3BkZXZzLW15LnNoYXJlcG9pbnQuY29tQDM0ZTg4NTQ3LTg4NjItNDg4Ny05NWYzLTgzOWJlNzkyZDM4NCIsImlzcyI6IjAwMDAwMDAzLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMEA5MDE0MDEyMi04NTE2LTExZTEtOGVmZi00OTMwNDkyNDAxOWIiLCJuYmYiOiIxNDEzMzE0NDY2IiwiZXhwIjoiMTQxMzM1MDQ2NiIsIm5hbWVpZCI6IjAjLmZ8bWVtYmVyc2hpcHxhZG1pbkBzamtwZGV2cy5vbm1pY3Jvc29mdC5jb20iLCJuaWkiOiJtaWNyb3NvZnQuc2hhcmVwb2ludCIsImlzdXNlciI6InRydWUiLCJjYWNoZWtleSI6IjBoLmZ8bWVtYmVyc2hpcHwxMDAzN2ZmZTg5ZWRmZmM3QGxpdmUuY29tIiwiaXNsb29wYmFjayI6IlRydWUiLCJhcHBjdHgiOiIwZThjZDk5ZWI3NzA0NjFjOWQxZDAzZTRhNTViYjhhYztNTUlvZFR1azJ4WnA0U0xSMlJUN3FPR3EzbE09O0RlZmF1bHQ7OzdGRkZGRkZGRkZGRkZGRkY7VHJ1ZTsifQ%2EMAfAMGPnwHDpi4mwNqoDg%2DlaXLW9HxKl2kEEw4fJkWs5B%5Fjd3wtj%2DFmIjytiDjB0kCsuv07NiKAYcyfwzGJzz8SFKoaCGHBGcRL3dooxJRZRZp%5FAOn5%5FMoMmG7bD239ZjizVdkJI5l2GUdzPq7CD52t535RtEUawDa7bnQrWfRQNdYcR1%5FlI5Ide04BifHMFEQ1gz9gsKnPQXDnGx%5FhizReK07q6e4w1YQMTMhTk7bjXHj9ZxwEU8Xa1usec8aQF2QhlXAKzgD7vw1PnqD%5FI46WnR69gilFJSwogoAQwkdnFdJz79EONc2GjqCcp%5Fh0VkyHxGJAmnsPq9aQFgHxpIA&access_token_ttl=1413350466258&z=%2522%257B0E8CD99E%252DB770%252D461C%252D9D1D%252D03E4A55BB8AC%257D%252C1%2522&type=png&o15=1&ui=en-US&PdfMode=1"
//var owaDocx = https://euc-word-view.officeapps.live.com/wv/docdatahandler.ashx?WOPIsrc=https%3A%2F%2Fsjkpdevs%2Dmy%2Esharepoint%2Ecom%2Fpersonal%2Fadmin%5Fsjkpdevs%5Fonmicrosoft%5Fcom%2F%5Fvti%5Fbin%2Fwopi%2Eashx%2Ffiles%2F718a86311f9a4b0e86a50ba9cc41eeb9&access_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjdrNTBKdl9vYVloQ3lHcDBOVUlQdkFYQ3l0TSJ9%2EeyJhdWQiOiJ3b3BpL3Nqa3BkZXZzLW15LnNoYXJlcG9pbnQuY29tQDM0ZTg4NTQ3LTg4NjItNDg4Ny05NWYzLTgzOWJlNzkyZDM4NCIsImlzcyI6IjAwMDAwMDAzLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMEA5MDE0MDEyMi04NTE2LTExZTEtOGVmZi00OTMwNDkyNDAxOWIiLCJuYmYiOiIxNDEzMzIyMjkwIiwiZXhwIjoiMTQxMzM1ODI5MCIsIm5hbWVpZCI6IjAjLmZ8bWVtYmVyc2hpcHxhZG1pbkBzamtwZGV2cy5vbm1pY3Jvc29mdC5jb20iLCJuaWkiOiJtaWNyb3NvZnQuc2hhcmVwb2ludCIsImlzdXNlciI6InRydWUiLCJjYWNoZWtleSI6IjBoLmZ8bWVtYmVyc2hpcHwxMDAzN2ZmZTg5ZWRmZmM3QGxpdmUuY29tIiwiaXNsb29wYmFjayI6IlRydWUiLCJhcHBjdHgiOiI3MThhODYzMTFmOWE0YjBlODZhNTBiYTljYzQxZWViOTtNTUlvZFR1azJ4WnA0U0xSMlJUN3FPR3EzbE09O0RlZmF1bHQ7OzdGRkZGRkZGRkZGRkZGRkY7VHJ1ZTsifQ%2EZ1UgMnYBgDNo%5FjFFdAIUrZ%5FwBA%2DF%2D2%2DNjg%5FNQk7PrdnF%5FGOCoX7s7B2E2CVnhhqOuBj%2DHB8kB2Mnmhw%5Fm95EXzH18eQ%5F6UN9ZWflEcKoSAcIh7FO0MHYddl3u%5FNzaWD1gxIr2HaUexEqMKmILjk7yEDjFDe7XlsFucIVTeyw9SHNtlovgNLZWNj7O%2DMJ5%5Ft3z5zoSr%2DcHg%2DqkyCcRwNPSfTeOWWUbAr9QMLz96DFKbxunpLN9fNO5x8tcYsw%5FFm1wvm7AA66A3ysRScZDd3eo3zzbz8ko0Ch3unV43fQiuRAMZd85Sn9gOx8riZDVIwZkia5tLGX9w4qAdEBDrkviA&access_token_ttl=1413358290958&z=%2522%257B718A8631%252D1F9A%252D4B0E%252D86A5%252D0BA9CC41EEB9%257D%252C4%2522&type=png&o15=1&ui=en-US
var owa = "/wv/docdatahandler.ashx?WOPIsrc={4}%5Fvti%5Fbin%2Fwopi%2Eashx%2Ffiles%2F{0}&access_token={1}&access_token_ttl={2}&type=png&o15=1&ui=en-US"

function encode(str) {
    str = str.replace(/_/g, '%5F');
    str = str.replace(/-/g, '%2D');
    str = str.replace(/\./g, '%2E');
    return str;
};
var progress = null;
function logProgress(str)
{
    progress.empty();
    progress.append(str);
}

$(document).ready(function () {
    progress = $('#progress');
    $('#btn').click(function () {
        var webUrl = $('#webUrl').val();
        $.ajax({
            url: webUrl + getFiles.replace('{0}',$('#listTitle').val()),
            type: "GET",
            headers: { Accept: 'application/json;odata=verbose' },
            success: function (res) {
                $.each(res.d.results, function (i, o) {
                    if (typeof (o.File.Name) !== 'undefined') {
                        console.log(o.File.UniqueId + " " + o.File.Name);
                        logProgress(o.File.Name);
                        $.get(webUrl + wopiframe.replace('{0}', '{' + o.File.UniqueId + '}')).done(function (res) {
                            //console.log(res);
                            $html = $(res);
                            var accessToken = $('input[name="access_token"]', $html).val();
                            var accessTokenTTL = $('input[name="access_token_ttl"]', $html).val();

                            var owaUrl = $('#owaUrl').val() + owa.replace("{0}", o.File.UniqueId.replace(/-/g, ''));
                            owaUrl = owaUrl.replace("{1}", encode(accessToken));
                            owaUrl = owaUrl.replace("{2}", accessTokenTTL);
                            owaUrl = owaUrl.replace("{4}", encode(encodeURIComponent(webUrl)))
                            if (o.File.Name.indexOf('.pdf') > 0) {
                                owaUrl += "&PdfMode=1";
                            }
                            console.log(owaUrl);


                            $.get(owaUrl).done(function (resw) {
                                //console.log(resw);
                                logProgress('Calling owa for ' + o.File.Name);
                                //console.log(new XMLSerializer().serializeToString(resw.documentElement));
                                //console.log();
                                try {
                                    var noPages = resw.getElementsByTagName("pageset")[0].getAttribute("count");
                                    $('#pages').append('<tr><td>' + o.File.Name + "</td><td> " + noPages + '</td></tr>');
                                }
                                catch (e) {
                                    console.log(new XMLSerializer().serializeToString(resw.documentElement));
                                    console.error(e);
                                }
                                finally
                                {
                                    logProgress('Done.')
                                }
                            });
                        });
                    }
                });

            }
        });
    });
});


// When the popup HTML has loaded
window.addEventListener('load', function (evt) {
    // Get the event page
    chrome.runtime.getBackgroundPage(function (eventPage) {
        // Call the getPageInfo function in the event page, passing in 
        // our onPageDetailsReceived function as the callback. This injects 
        // content.js into the current tab's HTML
        eventPage.getPageDetails(onPageDetailsReceived);
    });
});