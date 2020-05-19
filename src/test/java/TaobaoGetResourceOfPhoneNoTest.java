import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.utils.URLEncodedUtils;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.util.EntityUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.json.JSONObject;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;
import tools.ReadExcel;
import tools.ReadProperty;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.ResourceBundle;

public class TaobaoGetResourceOfPhoneNoTest {

    @DataProvider
    public Object[][] providerCaseData(){

        ReadExcel excel = null;
        try {
            excel =  new ReadExcel("淘宝手机号归属查询案例.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        int endrownum = excel.getRowNum(excel.getSheet("sheet1"));
        return excel.getBatchValues("sheet1",2,endrownum,2,3);

    }
    @Test(dataProvider = "providerCaseData" )
    public  void getPhonenumLocation(String telnum,String location) throws IOException {
        String url = ReadProperty.readValue("url");
//        System.out.println(url);
        //获得url地址
//        Assert.assertEquals(url,"https://tcc.taobao.com/cc/json/mobile_tel_segment.htm");

        //创建get请求需要的参数
        List<BasicNameValuePair> params =  new ArrayList<BasicNameValuePair>();
        BasicNameValuePair tel =  new BasicNameValuePair("tel",telnum);
        params.add(tel);
        //将参数转换为uri格式
        String uri = URLEncodedUtils.format(params,"UTF-8");
        String geturl =  url+"?".concat(uri);
        //创建get请求
        HttpGet get = new HttpGet(geturl);

        //创建客户端
        CloseableHttpClient client = HttpClients.createDefault();
        //客户端执行get请求
        HttpResponse response =  client.execute(get);
        //判断get请求返回码是否是200
        Assert.assertEquals(200,response.getStatusLine().getStatusCode());

        //获得response的返回实体
        HttpEntity httpentity =  response.getEntity();
        //将返回实体转换为String类型
        String  StringResponse = EntityUtils.toString(httpentity);
        //判断归属地，用字符串的方式判断包含
//        Assert.assertTrue(StringResponse.contains("北京联通"));

        //通过转换json的方式判断

        String [] a  = StringResponse.split("=");
        String strresponse =  a[1];
//        System.out.println(strresponse);
        JSONObject jsresponse = new JSONObject(strresponse);
//        System.out.println(jsresponse);
        Assert.assertEquals(location,String.valueOf(jsresponse.get("carrier")));
    }

    @DataProvider
    public Object[][] dataProvideTest1(){
//        Object[][] club1 = new Object[2][3];
//        club1=new Object[][]{{"1","2","3"},{"a","b","c"}};
        Object[][] club1 = new Object[2][3];
        club1[0] =  new Object[]{1,2,3};
        club1[1] =  new Object[]{4,5,6};






        for(int i=0;i<club1.length;i++){
            for(int j = 0;j<club1[i].length;j++){
                if(j<club1[i].length-1){
                    System.out.print(club1[i][j]);
                }else if(j==club1[i].length-1){
                    System.out.println(club1[i][j]);
                }

            }
        }

        return club1;
    }

    @Test(dataProvider = "dataProvideTest1")
    public void arrayTest(int a,int b,int c){
        System.out.println("test case start");
        System.out.println(a+b+c);
        System.out.println("test case end");
    }




}
