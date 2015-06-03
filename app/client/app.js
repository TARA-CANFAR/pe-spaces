Template.body.events({
    'change #file': function (event) {
        var reader = new FileReader(), 
            file = event.target.files[0];

        reader.onload = function (e) {
            var binary = e.target.result,
                workbook = XLSX.read(binary, {type: 'binary', cellStyles: true}),
                data = workbook.Sheets["Data"],
                output = workbook.Sheets["Output"],
                groups = [], processed = {}, rowPosition = 2, termColumns, i;
            
            
            // populate groups object from data worksheet
            for (z in data) {
                // all keys that do not begin with "!" correspond to cell addresses
                if (z[0] === '!') continue;
                
                var value = data[z].v.trim(), 
                    groupIds, genericInfo, column, year, classRange, term, location, i, name; 

                // select cells containing group data, given they all start with char 'Y'
                if (value.charAt(0) === "Y") { 
               
                    // create array of group ids, given they are separated by commas
                    groupIds = value.split(',')

                    // extract generic information from first group id
                    genericInfo = groupIds[0].split('_');
                    groupIds[0] = genericInfo[2]; /* i.e the specific info */

                    // assign generic information to separate vars
                    year = genericInfo[0].substr(1).trim();
                    classRange = genericInfo[1].trim();

                    // find column by removing number from cell address
                    column = z.replace(/\d+/g, '');

                    // scroll up the worksheet to identify location & term of groups
                    location = data[column + '4'].v.trim();
                    term = data[column + '1'].v.trim();

                    // add entry for each group above year 6
                    if (Number(year) > 6) {
                        for (i = 0; i < groupIds.length; i += 1) {
                            name = year + classRange + ' ' + groupIds[i].trim();
                            groups.push({name: name, location: location, term: term, 
                                         year: year, classRange: classRange, id: groupIds[i]});
                        }
                    }
                }
            }

            groups.sort(
                // sort groups by year
                firstBy(function (a, b) { return (a.year - b.year); })
                // sort groups by class range
                .thenBy(function (a, b) {
                    var nameA = a.classRange.toLowerCase(), nameB=b.classRange.toLowerCase()
                    if (nameA < nameB)
                        return -1 
                    if (nameA > nameB)
                        return 1
                    return 0
                })
                // sort groups by id
                .thenBy(function (a, b) {
                    var nameA = a.id.toLowerCase(), nameB=b.id.toLowerCase()
                    if (nameA < nameB)
                        return -1 
                    if (nameA > nameB)
                        return 1
                    return 0
                })
            );
            
            // clear previous data in output worksheet
            for (z in output) {
                // all keys that do not begin with "!" correspond to cell addresses
                if (z[0] === '!') continue;

                // start clearing at third row 
                if (Number(z.replace(/\D/g,'')) > 2) {
                    output[z].v = " ";
                }
            }
            
            // write groups into document
            for (i = 0; i < groups.length; i += 1) {
                var name = groups[i].name, term = groups[i].term, location = groups[i].location, 
                    targetCell;
                
                // check if group has been processed
                if (!processed[name]) {
                    rowPosition += 1;

                    // write new group names into new rows
                    output["A" + rowPosition].v = name;

                    // save row number
                    processed[name] = rowPosition;
                }
                

                // create mapping of terms and columns
                termColumns = {"1A": "B", "1B": "C", "2A": "D", 
                               "2B": "E", "3A": "F", "3B": "G"};
                
                // locate target cell based on term and rowPosition
                targetCell = output[termColumns[term] + processed[name]].v;
                
                // insert location into target cell
                if (targetCell.length > 1) {
                    output[termColumns[term] + processed[name]].v = targetCell + " / " + location;
                } else {
                    output[termColumns[term] + processed[name]].v = location;
                }
            }

            var wbout = XLSX.write(workbook, {type:'binary'});

            function s2ab(s) {
                var buf = new ArrayBuffer(s.length);
                var view = new Uint8Array(buf);
                for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                return buf;
            }
             
            /* the saveAs call downloads a file on the local machine */
            saveAs(new Blob([s2ab(wbout)],{type:""}), "output.xlsx")
        };

        reader.readAsBinaryString(file);
    }
})