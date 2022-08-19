const express=require("express");
const app=express();
const body_parse=require("body-parser");
const xlsx=require('node-xlsx');
const { resolveInclude }=require("ejs");
app.listen(3000);
app.use(body_parse.urlencoded({extended: true}));
app.use(express.static(__dirname+"/public"));
const multer =require('multer');
const fs=require('fs');
const path=require('path');
const storage=multer.diskStorage({
    destination:(req,file,cb)=>{
        cb(null,'uploads');
    },
    filename:(req,file,cb)=>{
        cb(null,file.fieldname+'-'+Date.now())
    }
});
const upload=multer({storage:storage});

app.get("/",(req,resp)=>{
    resp.sendFile(__dirname+"/public/html/upload.html");
});

function cmp_obj(a,b){
    if(a.product_line>b.product_line){
        return 1;
    }
    else if(a.product_line<b.product_line){
        return -1;
    }
    else{
        return 0;
    }
}
function process_data(data){
    let arr=[];
    let col_name=data[0];
    let p_pos=col_name.indexOf('PRODUCTLINE');
    let t_pos=col_name.indexOf('TERRITORY');
    let q_pos=col_name.indexOf('QUANTITYORDERED');
    let sv_pos=col_name.indexOf('SALES');
    let y_pos=col_name.indexOf('YEAR_ID');
    for(let i=1;i<data.length;i++){
        if(data[i][y_pos]==2004){
            let t_obj={
                product_line:data[i][p_pos],
                territory:data[i][t_pos],
                quantity:data[i][q_pos],
                salesvalue:data[i][sv_pos],
                year:data[i][y_pos]
            };
            arr.push(t_obj);
        }
    }
    arr.sort(cmp_obj);
    let ans_arr=[];
    ans_arr.push(['product-line','territory','quantity','sales']);
    for(let i=0;i<arr.length;i++){
        let t_arr=[];
        t_arr.push(arr[i].product_line);
        t_arr.push(arr[i].territory);
        t_arr.push(arr[i].quantity);
        t_arr.push(arr[i].salesvalue);
        ans_arr.push(t_arr);
    }
    return ans_arr;
}

app.post("/file",upload.single("excel_file"),(req,resp)=>{
    const file=xlsx.parse(path.join(__dirname+'/uploads/'+req.file.filename));
    const update_data=process_data(file[0].data);
    const buffer = xlsx.build([{name:'sheet1',data:update_data}]);
    const wpath='./public/exfile/'+req.file.filename+'.xlsx';
    fs.writeFileSync(wpath,buffer,{'flag':'w'});
    resp.render("download.ejs",{data_arr:update_data,action_value:"/downloadf/"+req.file.filename+'.xlsx'});
});

app.get("/downloadf/:filename",(req,resp)=>{
    resp.download('./public/exfile/'+req.params.filename);
});