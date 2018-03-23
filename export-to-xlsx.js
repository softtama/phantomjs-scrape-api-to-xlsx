/**
 * Export API data to XLSX file for various purposes (e.g. report, import product, etc.)
 * @version 1.0         First iteration (Mar 2018).
 * @author Rizki Pratama
 */

var Products = [], Indices = [],
    
    page    = require('webpage').create(),
    system  = require('system'),
    fs      = require('fs'),
    Q       = require('q'),
    XLSX    = require('XLSX'),
    
	apiUrl  = 'https://demo.whoisrizkipratama.net/phantomjs-scrape-api-to-xlsx-api-sample',
    
    rootPath            = 'D:/Repos/',
    subDirPath          = 'phantomjs-scrape-api-to-xlsx/',
    productIndicesPath  = rootPath + subDirPath + 'product-data-indices.txt',
    fileName            = 'product-report.xlsx',
    
    maxProductsPerSheet             = 100,  // Create new sheet after 100 product added
    maxProductsPerSheetIncrement    = 100;
    
/* Page settings */
page.onConsoleMessage = function (msg) { console.log(msg); };
page.settings.resourceTimeout = 600000; // 600 seconds = 10 mins.
/* END Page settings */

/* Helper functions */
/**
 * Determine whether an array contains a value.
 * https://stackoverflow.com/a/1181586
 */
var inArray = function (needle) {
    // Per spec, the way to identify NaN is that it is not equal to itself
    var findNaN = needle !== needle;
    var indexOf;
    
    if (!findNaN && typeof Array.prototype.indexOf === 'function') {
        indexOf = Array.prototype.indexOf;
    } else {
        indexOf = function(needle) {
            var i = -1, index = -1;

            for (i = 0; i < this.length; i++) {
                var item = this[i];

                if ((findNaN && item !== item) || item === needle) {
                    index = i;
                    break;
                }
            }

            return index;
        };
    }

    return indexOf.call(this, needle) > -1;
};

/**
 * Read a local file line by line using PhantomJS/Javascript.
 * https://stackoverflow.com/a/11755216
 */
var readLinesToArray = function (file_path) {
    var lines_array = [],
        content = '',
        f = null,
        lines = null,
        eol = system.os.name == 'windows' ? "\r\n" : "\n";
		
    try {
        f = fs.open(file_path, 'r');
        content = f.read();
    } catch (e) {
        console.log(e);
    }

    if (f) {
        f.close();
    }

    if (content) {
        lines = content.split(eol);
        for (var i = 0, len = lines.length; i < len; i++) {
            lines_array.push(lines[i]);
        }
    }
	
    return lines_array;
};
/* END Helper functions */

/**
 * Retrieve product data and store it to Product global variable.
 * @since 1.0
 */
var RetrieveProduct = function () {
    /**
     * Retrieve product data from API and filter it by product IDs read from a text file.
     * @return Q.Promise	Resolved when the data is successfully read and stored.
     * @since 1.0
     */
    this.readIndices = function () {
        var defer = Q.defer();

        console.log('==> Reading product IDs from text file...');
        Indices = readLinesToArray(productIndicesPath);
        defer.resolve(Indices);

        return defer.promise;
    };
	
    /**
     * Read product data from an API URL, filter it based on Indices, and store it to Product variable.
     * @return Q.Promise	Resolved when the product data is successfully retrieved,
     *                      rejected when the API URL is failed to open.
     * @since 1.0
     */
    this.readProductAPI = function () {
        var defer = Q.defer();

        console.log('==> Reading and filtering Product data from API URL...');

        page.open(apiUrl, function (status) {
            if (status !== 'success') {
                console.log('==> Error opening API page.');

                setTimeout(function () {
                    defer.reject('==> Failed to open API page.');
                }, 5000);
            } else {
                Products = page.evaluate(function (Indices, inArray) {
                    var element = document.getElementsByTagName('pre')[0],
                        products = (JSON.parse(element.textContent)).product,
                        products_aoa = [];

                    for (var i = 0; i < products.length; i++) {
                        // Filter based on IDs
                        if (!inArray.call(Indices, products[i].id.toString())) {
                            continue;
                        }
						
                        var product = [
                            products[i].id,
                            products[i].name,
                            products[i].category,
                            products[i].price,
                            products[i].weight,
                            products[i].description,
                            products[i].etalase,
                            products[i].condition
                        ];
						
                        // Continuing... push available image URLs
                        for (var img_index = 0; img_index < products[i].images.length; img_index++) {
                            product.push(products[i].images[img_index]);
                        }
                        
                        // Continuing... push available video URLs
                        for (var vid_index = 0; vid_index < products[i].videos.length; vid_index++) {
                            product.push(products[i].videos[vid_index]);
                        }
						
                        // Push to index based on Product ID
                        products_aoa.push(product);
                    }
					
                    // Return as array of array
                    return products_aoa;	
                }, Indices, inArray);
                
                defer.resolve(Products);
            }
        });

        return defer.promise;
    };
	
    /**
     * Start the chain processes.
     * @return Q.Promise
     * @since 1.0
     */
    this.start = function () {
        var defer = Q.defer(),
            retrieve_product = this;

        retrieve_product
            .readIndices()
            .then(function () {
                console.log('==> Done reading indices from text file!');
                console.log('==> - Number of indices          : ' + Indices.length);
                console.log('--------------------------------------------------------------');
                
                return retrieve_product.readProductAPI();
            })
            .done(function () {
                console.log('==> Done retrieving and filtering product data from API!');
                console.log('==> - Product length (filtered)  : ' + Products.length);
                console.log('==> - Max product added per sheet: ' + maxProductsPerSheet);
                console.log('--------------------------------------------------------------');
                
                defer.resolve();
            });
			
        return defer.promise;
    }
};

/**
 * Export retrieved product data to Excel file.
 * @since 1.0
 */
var ExportToExcelFile = function () {
    /**
     * Export product data to Excel worksheet in XLSX format.
     * @return Q.Promise	Resolved when the file is created,
     *                      rejected when the file is failed to create.
     * @since 1.0
     */
    this.exportToExcelFile = function () {
        var defer = Q.defer(),
            product_index = 0,
            wbbin;
		
        var header = [['ID', 'Name', 'Category', 'Price (in IDR)', 'Weight (in Gram)', 'Description', 
                'Etalase', 'Condition', 'Image 1', 'Image 2', 'Image 3', 'Image 4', 'Image 5', 
                'Video 1', 'Video 2', 'Video 3']],
            wscols = [
                {wch: 4},
                {wch: 24},
                {wch: 14},
                {wch: 10},
                {wch: 10},
                {wch: 20},
                {wch: 14},
                {wch: 8},
                {wch: 15},
                {wch: 15},
                {wch: 15},
                {wch: 15},
                {wch: 15},
                {wch: 15},
                {wch: 15},
                {wch: 15}
            ];
        
        console.log('==> Creating new workbook...');
        console.log('--------------------------------------------------------------');
        var wb = XLSX.utils.book_new();
		
        while (Products[product_index] !== undefined) {
            var products_for_xlsx = [];
            
            while (Products[product_index] !== undefined && product_index < maxProductsPerSheet) {
                products_for_xlsx.push(Products[product_index]);
                product_index++;
            }

            console.log('==> Creating new worksheet and adding product data with header to it...');
            var ws = XLSX.utils.aoa_to_sheet([['']]);
            ws['!cols'] = wscols;
            XLSX.utils.sheet_add_aoa(ws, header, { origin: { r: 0, c: 0 } });
            XLSX.utils.sheet_add_aoa(ws, products_for_xlsx, { origin: { r: 1, c: 0 } });

            var sheet_name = 'Product Report Sheet ' + (maxProductsPerSheet / 100);
            console.log('==> Appending worksheet "' + sheet_name + '" to the workbook...');
            XLSX.utils.book_append_sheet(wb, ws, sheet_name);
            
            maxProductsPerSheet += maxProductsPerSheetIncrement;
            console.log('--------------------------------------------------------------');
        }
		
        console.log('==> Creating binary data of workbook...');
        try {
            wbbin = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
        } catch (e) {
            defer.reject('==> Error creating workbook binary data: ' + e.toString());
        }
	
        console.log('==> Writing workbook as a new XLSX file: ' + fileName + ' ...');
        try {
            fs.write(fileName, wbbin, 'b');
        } catch (e) {
            defer.reject('==> Error writing file: ' + e.toString());
        }
		
        defer.resolve();

        return defer.promise;
    };

    /**
     * Start the chain processes.
     * @return Q.Promise
     * @since 1.0
     */
    this.start = function () {
        var defer = Q.defer(),
            export_to_xlsx = this;
			
        export_to_xlsx
            .exportToExcelFile()
            .done(function () {
                console.log('==> Done generating the Excel file!');
                defer.resolve();
            });

        return defer.promise;
    };
};

var App = function () {
    this.run = function () {
        app = this;
        var retrieve_product = new RetrieveProduct();
        
        retrieve_product
            .start()
            .then(function () {
                var export_to_xlsx = new ExportToExcelFile();
                export_to_xlsx.start();
            })
            .fail(function (reason) {
                console.log(reason + '\n==> Retrying from the beginning...');
                console.log('--------------------------------------------------------------');
                app.run();
            })
            .done(function () {
                console.log('==> Everything is done! You can close PhantomJS now...');
            });
    };
};

var app = new App();
app.run();
