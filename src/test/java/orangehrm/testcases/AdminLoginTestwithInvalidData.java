package orangehrm.testcases;

import java.io.IOException;

import org.testng.Assert;
import org.testng.annotations.Test;

import orangehrm.library.LoginPage;
import utils.AppUtils;
import utils.XLUtils;

public class AdminLoginTestwithInvalidData extends AppUtils 
{
	
	String datafile = "C:\\Selenium_8AM Batch\\OrangeHRM_DDT\\TestDataFiles\\Testdata.xlsx";
	String datasheet = "Login_InvalidData";

	@Test
	public void checkLogin() throws IOException
	{
		int rowcount = XLUtils.getRowCount(datafile, datasheet);
		
		String uid,pwd;
		LoginPage lp = new LoginPage();
		
		for(int i=1;i<=rowcount;i++)
		{
			uid = XLUtils.getStringData(datafile, datasheet, i, 0);
			pwd = XLUtils.getStringData(datafile, datasheet, i, 1);
			lp.login(uid, pwd);
			boolean res = lp.isErrMsgDisplayed();
			Assert.assertTrue(res);
			if(res)
			{
				XLUtils.setPass(datafile, datasheet, i, 2);
			}else
			{
				XLUtils.setFail(datafile, datasheet, i, 2);
			}			
		}
	}
	
	
}
