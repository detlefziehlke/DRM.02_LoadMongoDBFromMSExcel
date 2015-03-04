/**
 * Created by detlefziehlke on 04.03.15.
 */

"use strict";

var parseXlsx = require('excel'),
    q = require('q'),
    mongoose = require('mongoose'),
    fs = require('fs'),
    trans = {
      pathes: {
        excel: 'Projekte.xlsx', db: 'mongodb://localhost:27017/test'
      }
    };

fs.readdirSync(__dirname + '/models').forEach(function (filename) {
  if (~filename.indexOf('.js'))
    require(__dirname + '/models/' + filename);
});

var parseExcel = function (filename) {
  var deferred = q.defer();

  parseXlsx(trans.pathes.excel, function (err, data) {
    if (err) {
      deferred.reject(err);
      console.log('excel parsing error: ' + err);
    }
    else {
      deferred.resolve(data);
      console.log('excel parsing o.k.');
    }
  });

  return deferred.promise;
};

var openDb = function () {
  var deferred = q.defer();

  mongoose.connect(trans.pathes.db);
  mongoose.connection.on("open", function (err, con) {
    if (err) {
      deferred.reject(err);
      console.log('db connect error: ' + err);
    }
    else {
      deferred.resolve(con);
      console.log('db connect o.k.');
    }
  });

  return deferred.promise;
};

var dbStep0_1 = function () {
  var deferred = q.defer();

  mongoose.model('projects').collection.insert(trans.excelData, function (err, data) {
    if (err) {
      deferred.reject(err);
      console.log('dbStep0_1 error: ' + err);
    }
    else {
      deferred.resolve(data);
      console.log('dbStep0_1 o.k.');
    }
  });

  return deferred.promise;
};

var dbStep1_1 = function () {
  var query = mongoose.model('projects')
      .$where('this.name.length < 15')
      .where({area: {'$ne': 'harry'}})
      .sort('name -area')// alternative: .sort({name: 'desc', email:-1})
      .limit(20);
  return query.exec();
};

parseExcel(trans.pathes.excel)
    .then(function (data) {
      var dataObj = [];
      for (var i = 0; i < data.length; i++) {
        if (!data[i][0]) continue;
        dataObj.push({
          name: data[i][0], area: data[i][1]
        });
      }
      trans.excelData = dataObj;
      return openDb();
    })
    .then(function (data) {
      return dbStep0_1();
    })
    .then(function (data) {
      return dbStep1_1();
    })
    .then(function (data) {
      console.log('load o.k. ' + data.length + ' docs loaded');
      mongoose.connection.close();
    }, function (err) {
      mongoose.connection.close();
      console.log(err);
    })

;
