
class Orange_Locator:
    url = "https://opensource-demo.orangehrmlive.com/web/index.php/auth/login"
    user_name = "username"
    pass_name = "password"
    log_xpath = '//div[@class="oxd-form-actions orangehrm-login-action"]//button[@class="oxd-button oxd-button--medium oxd-button--main orangehrm-login-button"]'
    user_drop_xpath ='//span[@class="oxd-userdropdown-tab"]//p[@class ="oxd-userdropdown-name"]'
    logout_xpath = '//ul[@role="menu"]//li//a[@href="/web/index.php/auth/logout"]'
    dash_url = "https://opensource-demo.orangehrmlive.com/web/index.php/dashboard/index"
