/**
 * Created by detlefziehlke on 04.03.15.
 */

"use strict";

var parseXlsx = require('excel'),
    q = require('q');

var parseExcel = function (filename) {
  var deferred = q.defer();

  parseXlsx('Projekte.xlsx', function (err, data) {
    if (err) deferred.reject(err);
    else deferred.resolve(data);
  });

  return deferred.promise;
};

parseExcel('Projekte.xlsx').then(function (data) {
  var dataObj = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    dataObj.push({
      name: data[i][0], area: data[i][1]
    });
  }
  console.log(dataObj);
})

;
