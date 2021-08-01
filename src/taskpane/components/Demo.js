import React,{useState,useEffect} from 'react'
import Demo2 from './Demo2';

function Demo() {
    const [color, setcolor] = useState('black');
    const [val, setval] = useState(11);
    const [dis, setdis] = useState("none");


    const change_col=()=>{
        setcolor('green');
    }

    useEffect(() => {
        // Excel.run(function (context){
        //     // setval((prev)=>(prev+1));
        //     var sheet = context.workbook.worksheets.getItem('jhk');
        //     var range = sheet.getRange("C11");
        //     // var r2 = sheet.getRange("G17");
        //     // range.load('values');
        //     return context.sync()
        //     .then(function (){
        //         range.values="hello mr excel"
        //     // setval(vall);
        //     })
            
        //     range.format.autofitColumns();
           
        // });
        // var excel = new ActiveXObject("Excel.Application");
        // excel.Visible = true;
        // excel.Workbooks.Open("Invoice_jh_aus.xlsm");
       
    }, [])
    
    const open_demo2 = ()=>{
        setdis("block");
      }
  
    return (
        <>
       
       <button className="btn btn-primary" onClick={open_demo2}>Create Invoice</button>
        
        <div style={{'display': dis,"overflow":"hidden"}}>
        <Demo2/>
        </div>      
            
          
        </>
    )
}

export default Demo
