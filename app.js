const https = require('https');
const reader = require('xlsx')  
const file = reader.readFile('./Data.xlsx')
var alert = require('alert');



//Read data from the Master sheet
let readDataFromExcel = function(){
    return new Promise(async (resolve, reject) => {
        let data = []
                const temp = reader.utils.sheet_to_json(
                file.Sheets[file.SheetNames[0]])
                temp.forEach((res) => {
                    data.push(res)
                })
        resolve(data);
    });
}



//Invoke API Callout to COWIN API and fetch the results for the given code and date .
//Date is always TODAY's Date
let getAvailableSlots = function(pincode,todayDate){

            console.log('Please wait while we are searching the hospitals...');

            return new Promise(async (resolve, reject) => {
                       
                https.get('https://cdn-api.co-vin.in/api/v2/appointment/sessions/public/findByPin?pincode='+pincode+'&date='+todayDate, (resp) => {

                        let data = '';

                        // A chunk of data has been received.
                        resp.on('data', (chunk) => {
                            data += chunk;
                        });
                        // The whole response has been received. Print out the result.
                        resp.on('end', () => {
                            resolve(JSON.parse(data));
                        });

                        }).on("error", (err) => {
                            resolve(err.message);
                        console.log("Error: " + err.message);
                        });
            });
};


(async()=>{

        let todayDate = new Intl.DateTimeFormat('en-IN').format(new Date());

        let data = await readDataFromExcel();

        let availableHospitals = [];
        
        var isAnySlotAvailable = false;

        for(var slot of data){

                var availableSlot = await getAvailableSlots(slot.PINCODE,todayDate);
                for(var session of availableSlot.sessions){

                        var hospitalObj = {}
                        
                        var isAvailable = false;

                                if(session.available_capacity > 0){

                                            if(slot.TYPE_OF_DOSE == 'Dose1' && session.available_capacity_dose1 > 0 && slot.VACCINE_TYPE.includes(session.vaccine)
                                                && slot.MIN_AGE_LIMIT == session.min_age_limit){
                                                    if(slot.STATUS != 'AVAILABLE'){
                                                        isAnySlotAvailable = true;
                                                    }
                                                    slot.STATUS = 'AVAILABLE';
                                                    isAvailable = true;

                                            }

                                            else if(slot.TYPE_OF_DOSE == 'Dose2' && session.available_capacity_dose2 > 0 && slot.VACCINE_TYPE.includes(session.vaccine)
                                                && slot.MIN_AGE_LIMIT == session.min_age_limit){
                                                    if(slot.STATUS != 'AVAILABLE'){
                                                        isAnySlotAvailable = true;
                                                    }
                                                    slot.STATUS = 'AVAILABLE';
                                                    isAvailable = true;
                                            }  

                                }


                                if(isAvailable){
                                        hospitalObj.CENTER_ID = session.center_id;
                                        hospitalObj.HOSPITAL_NAME = session.name;
                                        hospitalObj.ADDRESS = session.address +','+session.block_name+','+session.district_name+','+session.state_name;
                                        hospitalObj.PINCODE = session.pincode;
                                        hospitalObj.VACCINE_TYPE = session.vaccine;
                                        hospitalObj.FEE = session.fee;

                                        if(slot.TYPE_OF_DOSE == 'Dose1'){
                                                hospitalObj.AVAILABLE_CAPACITY = session.available_capacity_dose1;
                                                hospitalObj.DOSE_TYPE = 'Dose1';
                                        }
                                        else{
                                                hospitalObj.AVAILABLE_CAPACITY = session.available_capacity_dose2;
                                                hospitalObj.DOSE_TYPE = 'Dose2';

                                        }

                                    
                                    availableHospitals.push(hospitalObj);
                                }
                                else{
                                    if(slot.STATUS != 'AVAILABLE')
                                            slot.STATUS = 'NOT AVAILABLE';
                                }
                }   

                if(availableSlot.sessions.length == 0){
                    slot.STATUS = 'NO VACCINATION CENTER FOUND';
                }


        }

        file.SheetNames = [];

        const ws = reader.utils.json_to_sheet(data);
  
        reader.utils.book_append_sheet(file,ws,"SearchPinCode")

        const ws1 = reader.utils.json_to_sheet(availableHospitals);
  
        reader.utils.book_append_sheet(file,ws1,"HOSPITAL_DETAILS")
        
        // Writing to our file
        reader.writeFile(file,'./Data.xlsx')

        console.log('Search Completed..');
        if(isAnySlotAvailable){
                    alert('Vaccination Slots are Available. Please check the Hospital Details Section of the Master Sheet');
        }


})();


