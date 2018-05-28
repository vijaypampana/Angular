export class Excelconfig {

    private inputpath = '';
    private fileName = '';
    private sheetName = 'Combined';
    private outputPath = '';
    private jsonStart = '{';
    private jsonEnd = '}';
    private keys: string[] = []; newKeys: string[] = [];
    private quotesCount: number = 0;

    //This method will return the filePath of the input file
    public getFilePath(): string {
        return this.inputpath + this.fileName;
    }

    //This method will return the sheetName
    public getSheetName(): string {
        return this.sheetName;
    }

  //This method will return the output path where the files needs to be saved
    public getOutputPath(): string {
        return this.outputPath;
    }

    //This method will read the excel data and combine the content delimited by &&&. The last column is the value of the field
    public readCellValue (i: any, sh: any, columnCount: any): string {
      let temp: string = '';
      for (let j = 1; j < columnCount; j++) {
        let str1 = '';
        str1 = sh.getRow(i).getCell(j).value;
        if ((j == (columnCount - 2)) && (str1 != null)) {
          const str2 = str1.split('.').join('&&&');
          temp = temp + '&&&' + str2;
        } else if (str1 != null) {
          if ( j == 1) {
            temp = str1;
          } else {
            temp = temp + '&&&' + str1;
          }
        }
      }

        return temp;
    }

    //This method will process the file
    public processFile(data: any) {
      let json: string = '';
      this.quotesCount = 1;
      for(let i =0; i < data.length; i++) {
        json += this.processRecord(data[i], i)
      }
      while(this.quotesCount > 0) {
        json += this.jsonEnd;
        this.quotesCount--;
      }
      this.saveFile(json);
    }

    //This method will process the record
    public processRecord(data: any, i: any): string {
      let json: string = '';
      //Removing the first column as it will be fileName
      data.shift();
      if(i == 0) {
        this.keys = [];
        json = this.readKeys(data);
      } else {
        this.newKeys = [];
        json = this.compareKeys(data);
      }
      return json;
    }

    //This method will process the first record of the file
    public readKeys(keysData: any): string  {
      let tempJson: string = this.jsonStart;
      let numberOfElement: number = keysData.length;
      let valuePair = this.getValuePair(keysData[numberOfElement-1], keysData[numberOfElement - 2]);
      keysData.forEach((value)=>this.keys.push(value));
      this.keys.pop();
      this.keys.pop();
      this.keys.forEach((value)=> {
        tempJson += '"' + value + '": { ';
        this.quotesCount++;
      });
      return tempJson+valuePair;
    }

    //This method will process the first record of the file
    public compareKeys(keysData: any): string {
      let jsonRow: string = '';
      let keysMatch: number = 0;
      let keysDidNotMatch: number = 0;
      let stop: boolean = false;
      let numberOfElements: number = keysData.length;
      let tempKeys: string[] = [];
      let valuePair = this.getValuePair(keysData[numberOfElements - 1], keysData[numberOfElements - 2]);
      keysData.forEach((value)=>this.newKeys.push(value));
      this.newKeys.pop();
      this.newKeys.pop();
      tempKeys = (this.newKeys.length >= this.keys.length)?this.keys:this.newKeys;
      for(let i = 0; i < tempKeys.length; i++) {
        if(!stop) {
          if(this.keys[i]==this.newKeys[i]) {
            keysMatch++;
          } else {
            stop = true;
          }
        }
      }
      if((keysMatch==this.newKeys.length) && (keysMatch==this.keys.length)) {
        jsonRow += ', ' + valuePair;
      } else {
        keysDidNotMatch = keysData.length - keysMatch;
        let tempKeysMatch: number = keysMatch;
        let tempKeysDidNotMatch: number = keysDidNotMatch;
        tempKeys = [];
        this.newKeys.forEach((value)=>tempKeys.push(value));
        while(tempKeysMatch > 0) {
          tempKeys.shift();
          tempKeysMatch--;
        }
        while(tempKeysDidNotMatch > 0) {
          jsonRow += this.jsonEnd;
          this.quotesCount--;
          tempKeysDidNotMatch--;
        }

        jsonRow += ', ';
        tempKeys.forEach((value)=> {
          jsonRow += '"' + value + '": {';
          this.quotesCount++
        });
        jsonRow += valuePair;
      }
      this.keys = [];
      this.newKeys.forEach((value)=>this.keys.push(value));
      this.newKeys =[];
      return jsonRow;
    }

    public getValuePair(data1:any, data2:any): string {
      let valuePair: string = '';
      if(isNaN(Number(data1))) {
        valuePair = '"' + data2 + '":"' + data1 + '"';
      } else {
        valuePair = '"' + data2 + '":' + data1;
      }
      return valuePair;
    }

    //This method will save the file and print the output to console.
    public saveFile(data:any) {
      console.log("----------------------------------------------------------------------------------------------------------------------------------------------------------------------");
      try {
        var test1: object = JSON.parse(data);
        console.log(JSON.stringify(test1, undefined, 2));
      } catch(e) {
        console.log(e);
        console.log(data);
      }
    }
}
