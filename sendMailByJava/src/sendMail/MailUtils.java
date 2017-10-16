package sendMail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Properties;
import java.util.regex.Pattern;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimeUtility;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
  
  
/** 
 * Java Mail 工具类 
 *  
 * @author XueQi 
 * @version 1.0 
 *  
 */  
public class MailUtils {  
    private static String host;  
    private static String username;  
    private static String password;  
    private static String from;  
    private static String nick;  
  
    static {  
        try {  
            // Test Data  
            host = "mail.capitalbiolife.com";  
            username = "report@capitalbiolife.com";  
            password = "bajd8888";
            from = "report@capitalbiolife.com";  
            nick = "report@capitalbiolife.com";  
            // nick + from 组成邮箱的发件人信息  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    }  
  
    /** 
     * 发送邮件 
     *  
     * @param to 
     *            收件人列表，以","分割 
     * @param subject 
     *            标题 
     * @param body 
     *            内容 
     * @param filepath 
     *            附件列表,无附件传递null 
     * @return 
     * @throws MessagingException 
     * @throws AddressException 
     * @throws UnsupportedEncodingException 
     */  
    public static boolean sendMail(String to, String subject, String body,  
            List<String> filepath){  
        // 参数修饰  
        if (body == null) {  
            body = "";  
        }  
        if (subject == null) {  
            subject = "无主题";  
        }  
        // 创建Properties对象  
        Properties props = System.getProperties();  
        // 创建信件服务器  
        props.put("mail.smtp.host", host);  
        props.put("mail.smtp.auth", "true"); // 通过验证  
        // 得到默认的对话对象  
        Session session = Session.getDefaultInstance(props, null);  
        // 创建一个消息，并初始化该消息的各项元素  
        MimeMessage msg = new MimeMessage(session);
        try {
        	nick = MimeUtility.encodeText(nick);  
            msg.setFrom(new InternetAddress(nick + "<" + from + ">"));  
            // 创建收件人列表  
            if (to != null && to.trim().length() > 0) {  
                String[] arr = to.split(",");  
                int receiverCount = arr.length;  
                if (receiverCount > 0) {  
                    InternetAddress[] address = new InternetAddress[receiverCount];  
                    for (int i = 0; i < receiverCount; i++) {  
                        address[i] = new InternetAddress(arr[i]);  
                    }  
                    msg.addRecipients(Message.RecipientType.TO, address);  
                    msg.setSubject(MimeUtility.encodeText(subject,MimeUtility.mimeCharset("gb2312"), null));  
                    // 后面的BodyPart将加入到此处创建的Multipart中  
                    Multipart mp = new MimeMultipart();  
                    // 附件操作  
                    if (filepath != null && filepath.size() > 0) {  
                        for (String filename : filepath) {  
                            MimeBodyPart mbp = new MimeBodyPart();  
                            // 得到数据源  
                            FileDataSource fds = new FileDataSource(filename);  
                            // 得到附件本身并至入BodyPart  
                            mbp.setDataHandler(new DataHandler(fds));  
                            // 得到文件名同样至入BodyPart  
                            mbp.setFileName(MimeUtility.encodeText(fds.getName()));  
                            mp.addBodyPart(mbp);  
                        }  
                        MimeBodyPart mbp = new MimeBodyPart();  
                        mbp.setText(body);  
                        mp.addBodyPart(mbp);  
                        // 移走集合中的所有元素  
                        filepath.clear();  
                        // Multipart加入到信件  
                        msg.setContent(mp, "text/html;charset=gbk");  
                    } else {  
                        // 设置邮件正文  
                        msg.setText(body);  
                    }  
                    // 设置信件头的发送日期  
                    msg.setSentDate(new Date());  
                    msg.saveChanges();  
                    // 发送信件  
                    Transport transport = session.getTransport("smtp");  
                    transport.connect(host, username, password);  
                    transport.sendMessage(msg,  
                            msg.getRecipients(Message.RecipientType.TO));  
                    transport.close();  
                    return true;  
                } else {  
                    System.out.println("None receiver!");  
                    return false;  
                }  
            } else {  
                System.out.println("None receiver!");  
                return false;  
            }  
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("发送失败！"); 
			return false;
		}
    }  
    
    public static HashMap<String, String> readXml(String fileName){  
        HashMap<String, String> map = new LinkedHashMap<String, String>();
    	boolean isE2007 = false;    //判断是否是excel2007格式  
        if(fileName.endsWith("xlsx"))  
            isE2007 = true;  
        try {  
            InputStream input = new FileInputStream(fileName);  //建立输入流  
            Workbook wb  = null;  
            //根据文件格式(2003或者2007)来初始化  
            if(isE2007)  
                wb = new XSSFWorkbook(input);  
            else  
                wb = new HSSFWorkbook(input);  
            Sheet sheet = wb.getSheetAt(0);     //获得第一个表单  
            Iterator<Row> rows = sheet.rowIterator(); //获得第一个表单的迭代器  
            while (rows.hasNext()) {  
                Row row = rows.next();  //获得行数据  
                //System.out.println("Row #" + row.getRowNum());  //获得行号从0开始  
                Iterator<Cell> cells = row.cellIterator();    //获得第一行的迭代器 
                StringBuffer sBuffer = new StringBuffer();
                while (cells.hasNext()) {  
                    Cell cell = cells.next();  
                    //System.out.println("Cell #" + cell.getColumnIndex());
                    switch (cell.getCellType()) {   //根据cell中的类型来输出数据  
                    case HSSFCell.CELL_TYPE_NUMERIC:  
                        //System.out.println(cell.getNumericCellValue());  
                        break;  
                    case HSSFCell.CELL_TYPE_STRING:  
                        //System.out.println(cell.getStringCellValue());
                        sBuffer.append(cell.getStringCellValue()+",");
                        break;  
                    case HSSFCell.CELL_TYPE_BOOLEAN:  
                        //System.out.println(cell.getBooleanCellValue());  
                        break;  
                    case HSSFCell.CELL_TYPE_FORMULA:  
                        //System.out.println(cell.getCellFormula());  
                        break;  
                    default:  
                        //System.out.println("unsuported sell type");  
                    break;  
                    }  
                }
                if(!"".equals(sBuffer.toString()) && null!=sBuffer){
                	String[] strArray = null;
                	if(sBuffer.toString().contains(":") && sBuffer.toString().contains("---")){
                		String newString = sBuffer.toString().substring(0,sBuffer.toString().indexOf("---"));
                		strArray = newString.split(":");
                	}else{
                		strArray = sBuffer.toString().split(",");
                	}
                	Pattern emailPattern = Pattern.compile("\\w+([-+.]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*");
                    if(strArray.length>2 && emailPattern.matcher(strArray[2]).find()){
                    	map.put(strArray[0],strArray[2]);
                    }else{
                    	map.put(strArray[0],strArray[1]);
                    }
                }
            }  
        } catch (IOException ex) {  
            ex.printStackTrace();  
        }
        return map;
    } 
  
    public static void main(String[] args) throws AddressException,  
            MessagingException, IOException {
    	/**
		 * ar0:输入excel路径(编号：邮箱)
		 * ar1:输出发送成功和发送失败日志的路径
		 * ar2:pdf结果报告路径
		 */
		String ar0 = args[0];//"C:\\Users\\Administrator\\Desktop\\mail\\test.xlsx";//
		String ar1 = args[1];//"C:\\Users\\Administrator\\Desktop\\mail";//args[1]
		String ar2 = args[2];//"C:\\Users\\Administrator\\Desktop\\mail\\芯健康-报告";//args[2]
    	System.setProperty("mail.mime.splitlongparameters","false");
    	StringBuffer sb = new StringBuffer();
    	sb.append("敬爱的晶典同仁：\r\n");
    	sb.append("    您好！\r\n");
    	sb.append("    首先，非常感谢您对员工健康活动的关注和高度的参与！\r\n");
    	sb.append("    您的PMRA芯片疾病和药物风险检测报告结果已经出来，请您查收附件。\r\n");
    	sb.append("    您有任何的问题和建议欢迎随时与我们联系！\r\n");
    	sb.append("    祝您：身体健康！工作顺利！\r\n    ");
           
    	HashMap<String, String> map = readXml(ar0);
    	System.out.println("-------------start-------------");
    	System.out.println("共读取到"+map.size()+"个‘编码：邮箱’，如下：");
    	for (java.util.Map.Entry<String, String> entry : map.entrySet()) {
    		System.out.println(entry.getKey()+":"+entry.getValue());
		}
    	StringBuffer sBufferfalied = new StringBuffer();
    	StringBuffer sBufferSuccess = new StringBuffer();
    	int failednum = 0;
    	int successnum = 0;
    	for (java.util.Map.Entry<String, String> entry : map.entrySet()) {
			//System.out.println(entry.getKey()+":"+entry.getValue());
			List<String> filepath = new ArrayList<>();  
	        filepath.add(ar2+"\\"+entry.getKey()+"\\1. PMRA芯片疾病风险基因检测报告.pdf");  
	        filepath.add(ar2+"\\"+entry.getKey()+"\\2. PMRA芯片药物风险基因检测报告.pdf");  
	        boolean flag = sendMail(entry.getValue(), "PMRA芯片基因检测报告",sb.toString(),filepath);  
	        if(flag==true){
	        	sBufferSuccess.append(entry.getKey()+":"+entry.getValue()+"---发送成功\r\n");
	            successnum++;
	        } else {
	        	sBufferfalied.append(entry.getKey()+":"+entry.getValue()+"---发送失败\r\n");
	        	failednum++;
			}
		}
    	String failedpath = ar1+"\\failed_"+System.currentTimeMillis()+".txt";
    	String successpath = ar1+"\\success_"+System.currentTimeMillis()+".txt";
        PrintWriter pw1 = new PrintWriter(new FileWriter(successpath));
        pw1.print(sBufferSuccess.toString());
        pw1.close();
        PrintWriter pw = new PrintWriter(new FileWriter(failedpath));
        pw.print(sBufferfalied.toString());
        pw.close();
        System.out.println("成功发送："+successnum+"个，记录存于"+successpath+"；\r\n失败发送："+failednum+"个，记录存于"+failedpath);
        System.out.println("-------------end-------------");
    }  
}  
