var XLSX = require('xlsx');
var JSZip = require('jszip');
var moment = require('moment');
var _ = require('lodash');
var fs = require('fs');

var USAGE = 'USAGE: node downloadSankey.js'
    + '[--dataFile=./sankey-dummy.json]'
    + '[--pdfFolder=./excel]'
    + '[--id=aa:bb:cc:dd:ee]'
    + '[--ip=91.0.0.1]'
    + '[--stime=2018-12-01]'
    + '[--etime=2019-01-01]'
    + '[--direction=inbound]'
    + '[--name=jun pc'
    ;

var opts = {
    'dataFile': 'sankey-dummy.json',
    'pdfFolder': './excel',
    'id': 'aa:bb:cc:dd:ee',
    'ip': '91.0.0.1',
    'stime': '2018-12-01',
    'etime': '2019-01-01',
    'direction': 'inbound',
    'name': 'jun pc'
};

var downloadFileCache = [];

/**
 * Read the command line arguments, validate, and map onto the opts object
 */
function readArgs() {
    var paramArgs = process.argv.slice(0,10);
    paramArgs.forEach(function(arg, i){
        var matchArray = arg.match(/--(.*)=(.*)/);
        if(matchArray && matchArray.length === 3){
            opts[matchArray[1]] = matchArray[2];
        }
    });
}

/**
 * 
 * @param data
 * @returns {*}
 */
function downloadToSpreadSheet(data, direction) {
    data = data || {};
    var spreadsheetData = {};
    var network_connectivity = data.network_connectivity || {};

    spreadsheetData.geo_ip_mapping = network_connectivity.geo_ip_mapping || [];
    spreadsheetData.ip_app_mapping = network_connectivity.ip_app_mapping || [];
    spreadsheetData.ip_category_mapping = data.ip_category_mapping || [];

    var wb = {
        SheetNames: ['Device Data', 'All Traffic Data', 'Geo Location to IP and Users', 'IP Users to Applications', 'IP Detail'],
        Sheets: {}
    };

    var snapShotMeta = opts;

    var metaData = {
        Sheets: [
            {
                header: ['Device Name', 'IP', 'MAC Address', 'Traffic Start Time', 'Traffic End Time', 'Traffic Direction'],
                arrayFormatData: function(){
                var ret = [{
                    'Device Name': snapShotMeta.name,
                    'IP': snapShotMeta.ip,
                    'MAC Address': snapShotMeta.id,
                    'Traffic Start Time': snapShotMeta.stime,
                    'Traffic End Time': snapShotMeta.etime,
                    'Traffic Direction':snapShotMeta.direction
                }];
    return ret;
}

},
    {
        header: ['Geo Location', 'IP', 'Alert Number from Geo to IP', 'Alert Severity from Geo to IP', 'Alert Ids from Geo to IP',
            'Application', 'Alert Number from IP to Application', 'Alert Severity from IP to Application', 'Alert Ids from IP to Application', 'Remote URL', 'Category', 'VPN', 'Malicious'],
            arrayFormatData: function(data) {
        var geo_ip_mapping = data.geo_ip_mapping;
        var ip_app_mapping = data.ip_app_mapping;
        var ip_category_mapping = data.ip_category_mapping;
        var ret = [];

        var ipMapToApp = {};
        var ipDetailMap = {};

        ip_app_mapping && ip_app_mapping.forEach(function(item){
            ipMapToApp[item.ip] = item;
    });

        ip_category_mapping && ip_category_mapping.forEach(function(item){
            ipDetailMap[item.name] = item;
    })

        geo_ip_mapping && geo_ip_mapping.forEach(function(item){
            var row = {};
        row['Geo Location'] = item.geo_location;
        item['ip_list'] && item['ip_list'].forEach(function(ipItem){
            row['IP'] = ipItem.name;
        var detail = ipDetailMap[ipItem.name];
        if(detail){
            row['Geo Location'] = detail.geolocation;
            row['Remote URL'] = detail.remoteURL;
            row['Category'] = detail.urlCat;
            row['VPN'] = detail.vpn ? 'vpn': '';
            row['Malicious'] = detail.malicious ? 'malicious': '';
        }

        if(ipItem.alert_metadata && ipItem.alert_metadata.count > 0){
            row['Alert Number from Geo to IP'] = ipItem.alert_metadata.count;
            row['Alert Severity from Geo to IP'] = ipItem.alert_metadata.severity;
            row['Alert Ids from Geo to IP'] = ipItem.alert_metadata.alert_list.join(',');
        }

        var app = ipMapToApp[ipItem.name];
        if(app){
            app.apps && app.apps.forEach(function(appItem){
                row['Application'] = appItem.name;
            if(appItem.alert_metadata && appItem.alert_metadata.count > 0){
                row['Alert Number from IP to Application'] = appItem.alert_metadata.count;
                row['Alert Severity from IP to Application'] = appItem.alert_metadata.severity;
                row['Alert Ids from IP to Application'] = appItem.alert_metadata.alert_list.join(',');
            }
            ret.push(_.extend({}, row));

        })
        }else {
            ret.push(_.extend({}, row));
        }
    })



    })

        return ret;
    }

    },
    {
        header: ['Geo Location', 'IP', 'Alert Number', 'Alert Severity', 'Alert Ids'],
            dataField: 'geo_ip_mapping',
        arrayFormatData: function(data){
        var ret = [];
        data.forEach(function(item){
            var row = {};
        row['Geo Location'] = item.geo_location;
        item['ip_list'] && item['ip_list'].forEach(function(ipItem){
            row['IP'] = ipItem.name;
        if(ipItem.alert_metadata && ipItem.alert_metadata.count > 0){
            row['Alert Number'] = ipItem.alert_metadata.count;
            row['Alert Severity'] = ipItem.alert_metadata.severity;
            row['Alert Ids'] = ipItem.alert_metadata.alert_list.join(',');
        }

        ret.push(_.extend({}, row));
    })



    })

        return ret;
    }

    },
    {
        header: ['IP', 'Application','Alert Number', 'Alert Severity', 'Alert Ids'],
            dataField: 'ip_app_mapping',
        arrayFormatData: function(data){
        var ret = [];
        data.forEach(function(item){
            var row = {};
        row['IP'] = item.ip;

        item.apps && item.apps.forEach(function(appItem){
            row['Application'] = appItem.name;
        if(appItem.portList && appItem.portList.length>0)row['Port List'] = appItem.portList.join(',');
        if(appItem.alert_metadata && appItem.alert_metadata.count > 0){
            row['Alert Number'] = appItem.alert_metadata.count;
            row['Alert Severity'] = appItem.alert_metadata.severity;
            row['Alert Ids'] = appItem.alert_metadata.alert_list.join(',');
        }

        ret.push(_.extend({}, row));
    })
    })

        return ret;
    }
    },
    {
        header: ['IP', 'Geo Location', 'Remote URL', 'Category', 'VPN', 'Malicious'],
            dataField: 'ip_category_mapping',
        arrayFormatData: function(data){
        var ret = [];
        data.forEach(function(item){
            var row = {};
        row['IP'] = item.name;
        row['Geo Location'] = item.geolocation;
        row['Remote URL'] = item.remoteURL;
        row['Category'] = item.urlCat;
        row['VPN'] = item.vpn ? 'vpn': '';
        row['Malicious'] = item.malicious ? 'malicious': '';
        ret.push(_.extend({}, row));
    })

        return ret;
    }
    }
]
};


    var sheets = metaData.Sheets;

    wb.SheetNames.forEach(function(name, index){
        var sheet = sheets[index];
    var header = sheet.header;
    var data = sheet.dataField ? spreadsheetData[sheet.dataField] : spreadsheetData;
    var rows = sheet.arrayFormatData.call(this, data);

    var ws = XLSX.utils.json_to_sheet(rows, {header: header})
    wb.Sheets[name] = ws;
});
    //function s2ab(s) {
    //    var buf = new ArrayBuffer(s.length);
    //    var view = new Uint8Array(buf);
    //    for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    //    return buf;
    //}
    //var zipblob = s2ab(XLSX.write(wb, {bookType:'xlsx', type:'binary'}));
    var zipblob = XLSX.write(wb, {bookType:'xlsx', type:'binary'});
    var fileTime = moment['utc'].call().format('MMMM-DD-YYYYTHH-MM');
    var fileName = opts.pdfFolder + '/zingbox_sankey_' + fileTime + snapShotMeta.name + '_' + direction + '_traffic.xlsx';


    //downloadFileCache.push({
    //    filename: fileName,
    //    content: zipblob
    //});

    if (!fs.existsSync(opts.pdfFolder)){
        fs.mkdirSync(opts.pdfFolder);
    }

    fs.writeFile(fileName, zipblob, function(err){
               if (err) throw err;
                console.log('The file has been saved!');
            });






}

//function combineToZip(zip, filename) {
//    downloadFileCache.forEach(function(item){
//        zip.file(item.filename, item.content, {binary: true});
//    })
//
//    zip.generateAsync({type:"nodebuffer"})
//        .then(function(content) {
//            // see FileSaver.js
//            fs.writeFile(filename, content, function(err){
//                if (err) throw err;
//                console.log('The file has been saved!');
//            });
//        });
//}


/**
 * Create one zip file based on one device json file
 */
function createXLSXFile() {
    var dataStr = fs.readFileSync(opts.dataFile);
    var dataJson = JSON.parse(dataStr);
    //var timestamp = new Date().getTime();

    //var directions = ['inbound','outbound', 'all'];
    //var zip = new JSZip();

    //todo: json file should have 3 direction data
    downloadToSpreadSheet(dataJson, opts.direction);

    //directions.forEach(function(direction){
    //    downloadToSpreadSheet(dataJson, direction);
    //})

    //var filename = 'sankey-' + snapShotMeta.name + '-' + timestamp + '.zip';
    //
    //combineToZip(zip, filename);


}


readArgs();
createXLSXFile();





