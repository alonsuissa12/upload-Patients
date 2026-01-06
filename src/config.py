

class Config(object):

    def __init__(self,model):
        debug = False
        self.first_name_col = 0
        self.last_name_col = 1
        self.id_col = 2
        self.date_col = 3
        self.receipt_col = 4

        self.XL_path = ""
        if model == "clalit":
            self.did_reported_col = 4
            self.did_file_upload_col = 5
            self.left_over_treatment_col = 6
            self.need_new_approval_col = 7
            self.error_col = 8
            self.system_message_col = 10
            self.new_approval_file_col = 10
            self.is_referral_uploaded_col = 9
            self.login_name = "sm81471"
            self.login_verification = "123"
            self.site_link = "https://portalsapakim.mushlam.clalit.co.il/Mushlam/Login.aspx?ReturnUrl=%2fMushlam"
            self.wait_time_limit = 200 # seconds
            self.model = "clalit"

        else: # model == "macabi"
            self.did_reported_col = 5
            self.login_id = "126280"
            self.login_password = ""
            self.provider_type = "5"
            self.provider_code = "24657"
            self.site_link = "https://wmsup.mac.org.il/mbills"
            self.left_over_treatment_col = 6
            self.need_new_approval_col = 7
            self.error_col = 8

            self.model = "macabi"


