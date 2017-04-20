// Copyright (c) 2017 Olivier Gérardin
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
//
//
// TrackAndTrace function for POST Luxembourg registered shipments.
//
// Provides a function "trackingStatus" you can use in an Google Sheets document to obtain the latest status of a registered shipment.
// The result is a semi-colon delimited string in the following format: "<date>;<time>;<place>;<status>". If your shipment reference
// is in a cell e.g. E2, you can use the formula "=trackingStatus(E2)" to retrieve the tracking status, and then use the SPLIT
// function to distribute the parts of the status in individual cells.
// 
// Unfortunately (afaik) POST Luxembourg doesn't provide a public API for this, so the status is obtained by scraping their "Track and Trace" 
// page at http://www.trackandtrace.lu
//
// I've implemented 2 distinct methods: one using direct HTTP POST on the URL and walking the resulting HTML through builtin services only,
// and another using YQL (Yahoo Query Language) to directly extract the relevant data from an XPath. Both works, the native one
// is probably preferable as it doesn't rely on an external API, but it uses a deprecated service for leniently parsing HTML
// that Google might decide to remove anytime.
//
// WARNINGS: 
// -might or might not work with shipments from other carriers than POST Luxembourg
// -not tested extensively, no sanity checks, USE AT YOUR OWN RISK!


// Main function
function trackingStatus(n) {
  // choose your implementation...
  return _trackNative(n);
  //return _trackYql(n);
}

// Native version: using only builtin GScript services
function _trackNative(n) {
  var url = "http://www.trackandtrace.lu/homepage.htm";
  var options = {
    "method": "post",
    "headers": {
    },
    "payload": {
      "numero": n
    }
  };
  var response = UrlFetchApp.fetch(url, options);
  var html = response.getContentText();

  // This doesn't work because the URL returns malformed HTML
  //var doc = XmlService.parse(html);  
  
  // So we use this trick to parse leniently: http://stackoverflow.com/questions/19455158/what-is-the-best-way-to-parse-html-in-google-apps-script
  var doc0 = Xml.parse(html, true); // careful: deprecated service
  var bodyXml = doc0.html.body.toXmlString();
  var doc = XmlService.parse(bodyXml);
  
  var root = doc.getRootElement();
  
  var el = getElementById(root, "tr" + n);
  var div = el.getChildren('div')[1].getChildren('div')[1];  
  
  var date = div.getChildren('div')[0].getChildren('p')[0].getText().trim();
  var time = div.getChildren('div')[1].getChildren('p')[0].getText().trim();
  var place = div.getChildren('div')[2].getChildren('p')[0].getText().trim();
  var status = div.getChildren('div')[3].getChildren('p')[0].getText().trim();
  
  return date + ";" + time + ";" + place + ";" + status;
}

// version using YQL through remote API to fetch document and extract XPath as JSON
function _trackYql(n) {
  // base URL for tracking; takes a post request with parameters like "numero=RRxxxxxxxxLU"
  // the result is HTML so we'll have to scrape it to extract the data
  var url = "http://www.trackandtrace.lu/homepage.htm";

  // This XPath extracts the first lin in the history table, which contains the latest status data  
  var xpath= "//*[@id=\"tr" + n + "\"]/div[2]/div[2]"

  // Little trick to use YQL with the results of a HTTP POST
  // Adapted from https://www.christianheilmann.com/2009/11/16/using-yql-to-read-html-from-a-document-that-requires-post-data
  var query= "use 'https://raw.githubusercontent.com/yql/yql-tables/master/data/htmlpost.xml' as htmlpost; select * from htmlpost where url='" + url + "' and postdata='numero=" + n + "' and xpath='" + xpath + "'";
  
  var yql   = "https://query.yahooapis.com/v1/public/yql?format=json&q=" + encodeURIComponent(query);
  
  var response = UrlFetchApp.fetch(yql);
  var json = JSON.parse(response.getContentText());
  
  var date = json.query.results.postresult.div.div[0].p;
  var time = json.query.results.postresult.div.div[1].p;
  var place = json.query.results.postresult.div.div[2].p;
  var status = json.query.results.postresult.div.div[3].p;
  
  //status contains UTF-8 interpreted as windows-125, this little trick fixes it
  status = decodeURIComponent(escape(status));
  
  return date + ";" + time + ";" + place + ";" + status;
}

function _trackTest() {
  //var res = _trackYql("RR050147454LU");
  var res = _trackNative("RR050147454LU");
  Logger.log(res);  
}

// From https://sites.google.com/site/scriptsexamples/learn-by-example/parsing-html
function getElementById(element, idToFind) {  
  var descendants = element.getDescendants();  
  for(i in descendants) {
    var elt = descendants[i].asElement();
    if( elt !=null) {
      var id = elt.getAttribute('id');
      if( id !=null && id.getValue()== idToFind) return elt;    
    }
  }
}
