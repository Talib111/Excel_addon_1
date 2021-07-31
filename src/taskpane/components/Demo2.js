import React,{useState} from 'react'
function Demo2() {

    const [inputval, setinputval] = useState("");
    const [item_no, setitemno] = useState([1]);
    //product array
  const [select_pro_input, setSelect_pro_input] = useState([]);
  //cost array
  const [select_cost_input, setSelect_cost_input] = useState([]);
  //qty array
  const [select_qty_input, setSelect_qty_input] = useState([]);
  //gst array
  const [select_gst_input, setSelect_gst_input] = useState([]);

  //single object of all array
  const [all, setall] = useState({});

    const [patient, setpatient] = useState('');
    const [address, setaddress] = useState('');
    const [phone, setphone] = useState('');


    //receingng select product value
  const select_change_pro = (e) => {
    var s_id = e.target.id;
    var s_val = e.target.value;

    setSelect_pro_input({ ...select_pro_input, [s_id]: s_val });

  };
  const select_change_cost = (e) => {
    var s_id = e.target.id;
    var s_val = e.target.value;

    setSelect_cost_input({ ...select_cost_input, [s_id]: s_val });
  };
  const select_change_qty = (e) => {
    var s_id = e.target.id;
    var s_val = e.target.value;

    setSelect_qty_input({ ...select_qty_input, [s_id]: s_val });
  };
  const select_change_gst = (e) => {
    var s_id = e.target.id;
    var s_val = e.target.value;

    setSelect_gst_input({ ...select_gst_input, [s_id]: s_val });
  };

  var pro_array;
  var cost_array;
  var qty_array;
  var gst_array;
  //merging in single array
  const merge_array = ()=>{
   pro_array = Object.values(select_pro_input);
   cost_array = Object.values(select_cost_input);
   qty_array = Object.values(select_qty_input);
   gst_array = Object.values(select_gst_input);

    var all_pro = [];

    for(var i=0;i<=item_no.length - 1;i++){
      //setstate was making proble in the loop so took simple varaible
    //  all_pro = {...all_pro,["row_"+i]:[pro_array[i],cost_array[i],qty_array[i],gst_array[i]]};
     all_pro = [...all_pro,["["+JSON.stringify(pro_array[i]),JSON.stringify(cost_array[i]),JSON.stringify(qty_array[i]),JSON.stringify(gst_array[i])+"]"]];
    }
    // document.getElementById('all').innerHTML=Object.values(all_pro);
    document.getElementById('all').innerHTML=all_pro;
    save_to_excel();
  }


    const add_item = () => {


      //====1==== collect item in item_array state
      // setitem_array()
        //=================input validation=================
        // var local_ppp = document.getElementById("product_"+item_no[item_no.length-1]).value;
        // var local_vvv = document.getElementById("value_"+item_no[item_no.length-1]).value;
        // console.log(local_ppp);
        
    //     if(local_ppp==""){
    //      setp_error('Enter products !')
    //    }
    //     //valid product entry
    //     else if(!product_include.includes(local_ppp)){
    //      setp_error("Enter valid product!");
    //    }
       
    //    else if(local_vvv==""){
    //      setp_error('');
    //      setv_error("Enter value !")
    //    }
      
    //    //=================input validation=================
    //   else{
        //  setp_error('');
        //  setv_error("");
        var last_em = item_no[item_no.length - 1];
        setitemno([...item_no, last_em + 1]);
    //    }
      };




    const save_to_excel = ()=>{
        Excel.run(function (context){
            // setval((prev)=>(prev+1));
            var sheet = context.workbook.worksheets.getItem('jhk');
            // var range = sheet.getRange("A10");
           
           
            return context.sync()
            .then(function (){
              for(var i=0;i<=item_no.length - 1;i++){
               var idn = i+15;
              sheet.getRange("A"+idn).values = pro_array[i];
              sheet.getRange("B"+idn).values = cost_array[i];
              sheet.getRange("C"+idn).values = qty_array[i];
              sheet.getRange("D"+idn).values = gst_array[i];
              }

              // sheet.getRange("A10").values = patient;
              // sheet.getRange("A11").values = address+" "+phone;
              // window.print();
              // Application.Dialogs(xlDialogPrint).Show;
              // sheet.xlDialogPrint();
              // sheet.getRange("A16").EntireRow.Insert;
            })
            
            // range.format.autofitColumns();
           
        });
    }

    

//     const save_to_excel = (all_pro)=>{
//       Excel.run(function (context) {
//     var sheet = context.workbook.worksheets.getItem("jhk");
//     sheet.getRange('A10').values = patient;
//     sheet.getRange('A11').values = address+" "+phone;
    
//     var expensesTable = sheet.tables.add("A15:D15", true /*hasHeaders*/);
//     expensesTable.name = "ExpensesTable3";

//     // expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];
//     expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    
//     // expensesTable.rows.add(null /*add rows to the end of the table*/, [
//     //     ["1/1/2017", "The Phone Company", "Communications", "$120"],
//     //     ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
//     //     ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
//     //     ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
//     //     ["1/11/2017", "Bellows College", "Education", "$350"],
//     //     ["1/15/2017", "Trey Research", "Other", "$135"],
//     //     ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
//     // ]);
//     expensesTable.rows.add(null /*add rows to the end of the table*/, [
//       all_pro
//   ]);

//     if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
//         sheet.getUsedRange().format.autofitColumns();
//         sheet.getUsedRange().format.autofitRows();
//     }
   
//     sheet.activate();

//     return context.sync();
// });
//     }
    return (
        <div>
            <h5 className="bg-sucess text-center">{JSON.stringify(Object.values(select_pro_input))}</h5>
            <h5 className="bg-sucess text-center">{JSON.stringify(select_cost_input)}</h5>
            <h5 className="bg-sucess text-center">{JSON.stringify(select_qty_input)}</h5>
            <h5 className="bg-sucess text-center">{JSON.stringify(select_gst_input)}</h5>
            <h5 className="bg-sucess text-center">{item_no}</h5>
            <h5 className="bg-sucess text-center" id="all">final array</h5>
         
           <div className="container bg-light pb-5">
              
                    <div className="row px-0">
                        <div className="col-4 px-0"><label htmlFor="patient">Patinet Name</label> <input id="patient" type="text" onChange={(e)=>{setpatient(e.target.value)}}  /></div>
                        <div className="col-4"><label htmlFor="address">Address</label><input id="address" type="text" onChange={(e)=>{setaddress(e.target.value)}}  /></div>
                        <div className="col-4"><label htmlFor="phone">Phone</label><input id="phone" type="text" onChange={(e)=>{setphone(e.target.value)}}  /></div>
                    </div>
                    <div className="row px-0 mt-3">
                        <div className="col-3 px-0"><h6>Product</h6></div>
                        <div className="col-3"><h6>Cost</h6></div>
                        <div className="col-3"><h6>Qty</h6></div>
                        <div className="col-3"><h6>Gst</h6></div>
                    </div>
                   {item_no.map((etm)=>(
                        <div className="row px-0 mt-2" id={"row_"+etm}>
                        <div className="col-3 px-0">
                        <input
                         style={{width: "100%",paddingLeft: '7px'}}
                         type="text"
                         list={"d_product_"+etm}
                         id={"product_"+etm}
                         onChange={select_change_pro}
                       />
                       <datalist id={"d_product_"+etm}>
                         <option value="Png">Png - 001</option>
                         <option value="Pnr">Pnr - 002</option>
                         <option value="Aj100">Aj100 - 003</option>
                         <option value="NBP">NBP - 004</option>
                         <option value="HC">HC - 005</option>
                         <option value="NGT">NGT - 006</option>
                         <option value="JKN">JKN - 007</option>
                         <option value="ANANDAM">ANANDAM -- 08</option>
                         <option value="MASIHI_P">MASIHI(P) -- </option>
                         <option value="SUGAR_P">SUGAR(P)</option>
                         <option value="MSG_P">MSG(P)</option>
                         <option value="MSG_G">MSG(G)</option>
                         <option value="DAMA_P">DAMA(P)</option>
                         <option value="KB100">KB100</option>
                        
                       </datalist>
                        </div>
                    
                   
                        <div className="col-3 px-2">
                        <input
                         style={{width: "100%",paddingLeft: '7px'}}
                         type="text"
                         list={"d_cost_"+etm}
                         id={"cost_"+etm}
                         onChange={select_change_cost}
                       />
                       <datalist id={"d_cost_"+etm}>
                         <option value="Png">Png - 001</option>
                         <option value="Pnr">Pnr - 002</option>
                         <option value="Aj100">Aj100 - 003</option>
                         <option value="NBP">NBP - 004</option>
                         <option value="HC">HC - 005</option>
                         <option value="NGT">NGT - 006</option>
                         <option value="JKN">JKN - 007</option>
                         <option value="ANANDAM">ANANDAM -- 08</option>
                         <option value="MASIHI_P">MASIHI(P) -- </option>
                         <option value="SUGAR_P">SUGAR(P)</option>
                         <option value="MSG_P">MSG(P)</option>
                         <option value="MSG_G">MSG(G)</option>
                         <option value="DAMA_P">DAMA(P)</option>
                         <option value="KB100">KB100</option>
                        
                       </datalist>
                        </div>
                   
                   
                        <div className="col-3 px-2">
                        <input
                         style={{width: "100%",paddingLeft: '7px'}}
                         type="text"
                         list={"d_qty_"+etm}
                         id={"qty_"+etm}
                         onChange={select_change_qty}

                       />
                       <datalist id={"d_qty_"+etm}>
                         <option value="Png">Png - 001</option>
                         <option value="Pnr">Pnr - 002</option>
                         <option value="Aj100">Aj100 - 003</option>
                         <option value="NBP">NBP - 004</option>
                         <option value="HC">HC - 005</option>
                         <option value="NGT">NGT - 006</option>
                         <option value="JKN">JKN - 007</option>
                         <option value="ANANDAM">ANANDAM -- 08</option>
                         <option value="MASIHI_P">MASIHI(P) -- </option>
                         <option value="SUGAR_P">SUGAR(P)</option>
                         <option value="MSG_P">MSG(P)</option>
                         <option value="MSG_G">MSG(G)</option>
                         <option value="DAMA_P">DAMA(P)</option>
                         <option value="KB100">KB100</option>
                        
                       </datalist>
                     
                    </div>
                  
                        <div className="col-2 px-0">
                        <input
                         style={{width: "100%",paddingLeft: '7px'}}
                         type="text"
                         list={"d_gst_"+etm}
                         id={"gst_"+etm}
                         onChange={select_change_gst}

                       />
                       <datalist id={"d_gst_"+etm}>
                         <option value="Png">Png - 001</option>
                         <option value="Pnr">Pnr - 002</option>
                         <option value="Aj100">Aj100 - 003</option>
                         <option value="NBP">NBP - 004</option>
                         <option value="HC">HC - 005</option>
                         <option value="NGT">NGT - 006</option>
                         <option value="JKN">JKN - 007</option>
                         <option value="ANANDAM">ANANDAM -- 08</option>
                         <option value="MASIHI_P">MASIHI(P) -- </option>
                         <option value="SUGAR_P">SUGAR(P)</option>
                         <option value="MSG_P">MSG(P)</option>
                         <option value="MSG_G">MSG(G)</option>
                         <option value="DAMA_P">DAMA(P)</option>
                         <option value="KB100">KB100</option>
                        
                       </datalist>
                        </div>
                        <div className="col-1"><div style={{'cursor':'pointer'}}>❌</div></div>
                        </div>
                   ))}
                  
              
               <button className="btn btn-sm btn-primary text-white mt-3 mb-3"  onClick={add_item}>➕ Add Product</button>
               <div className="row mt-5">
                 <div className="col-sm-4"> <label htmlFor="discount">Discount %</label> <input id="discount" type="text" value="0" /></div>
                 <div className="col-sm-6 offset-2"><label htmlFor="mode">Mode of Payment</label><div><select className="px-2" name="mode" id="mode"><option value="cash">CASH</option>
                 <option value="online">ONLINE</option></select></div></div>
               </div>
               <button className="btn text-white mt-5 px-5 py-1" style={{backgroundColor: 'green',fontWeight: 700}} onClick={merge_array}>Print</button>
           </div>
        </div>
    )
}

export default Demo2
