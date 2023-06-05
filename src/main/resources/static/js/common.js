/**

 *
 */

function shutdown() {
  var xhr = new XMLHttpRequest();
  xhr.open("POST", "/shutdown", true);
  xhr.send(null);
}