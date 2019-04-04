import pyodbc
import datetime
import ldap3


class HmailMailSync:
    '''
    Iss script only for MS LDAP
    '''
    def __init__(self, ldap_server, ldap_user, ldap_password, search_base):
        self.ldap_server = ldap_server
        self.ldap_user = ldap_user
        self.ldap_password = ldap_password
        self.search_base = search_base

    def get_ldap_users(self):

        ldap = ldap3.Connection(server=self.ldap_server, user=self.ldap_user,
                                password=self.ldap_password, auto_bind=True)
        # user, who got param in AD about: mail, department and sn
        ldap.search(search_base=self.search_base, search_filter='(&(mail=*)(sn=*)(department=*))',
                    attributes='sAMAccountName')
        _ = []
        for item in ldap.entries:
            _.append(str(item.sAMAccountName))
        _.sort()
        print(f'Count of existing LDAP users = {len(_)}')
        return set(_)

    def get_sql_users(self, cur):
        str_to_exec = "SELECT accountadusername FROM dbo.hm_accounts"
        cur.execute(str_to_exec)
        _ = []
        for item in cur.fetchall():
            # print(item[0])
            _.append(item[0])
        print(f'Count of users within SQL  = {len(_)}')
        return set(_)

    def get_id_user(self, user, cur):
        # Collect all SQL users only to get set difference
        str_to_exec = f"SELECT accountID FROM dbo.hm_accounts WHERE accountadusername='{user}'"
        cur.execute(str_to_exec)
        id_user = cur.fetchone()[0]
        print(f'{user} have ID = {id_user}')
        return id_user

    def create_mails_and_boxes(self, user, cur, server, con):
        now = str(datetime.datetime.now())
        print(f'Creating email_box for : {user}')
        # now[:-3] its correct time format on tested SQL server
        # There is two steps. 1) Create email account
        str_to_exec = "INSERT into dbo.hm_accounts(accountdomainid,accountadminlevel,accountaddress,accountpassword," \
                      "accountactive,accountisad,accountaddomain,accountadusername,accountmaxsize,accountvacationmessageon," \
                      "accountvacationmessage,accountvacationsubject,accountpwencryption,accountforwardenabled," \
                      "accountforwardaddress,accountforwardkeeporiginal,accountenablesignature,accountsignatureplaintext," \
                      "accountsignaturehtml,accountlastlogontime,accountvacationexpires,accountvacationexpiredate," \
                      "accountpersonfirstname,accountpersonlastname) " \
                      "values(1,0,'{}@{}','',1,1,'geo.kr','{}',0,0,'','',0,0,'',0,0,'','','{}',0,'{}','',''" \
                      ")".format(user, server, user, now[:-3], now[:-3])

        cur.execute(str_to_exec)
        con.commit()

        id_user = self.get_id_user(user, cur)
        # There is two steps. 2) Create email IMAP folder, even you don't use IMAP
        str_to_exec = "INSERT into dbo.hm_imapfolders (folderaccountid,	folderparentid,	foldername,	" \
                      "folderissubscribed,foldercreationtime, foldercurrentuid) " \
                      f"VALUES ({id_user}, -1, 'INBOX', 1, '{now[:-3]}', 0)"
        cur.execute(str_to_exec)
        con.commit()


def manual_update(class_that_update_hmail, sql_server, sql_user, sql_database_mail, sql_password):

    con = pyodbc.connect(f'DRIVER={{SQL Server}}; SERVER={sql_server}; DATABASE={sql_database_mail};'
                         f' UID={sql_user}; PWD={sql_password}')
    cur = con.cursor()
    user_with_nosql_mail = class_that_update_hmail.get_ldap_users() - class_that_update_hmail.get_sql_users(cur)
    print(f'Users without SQL mail {user_with_nosql_mail}')

    for _ in user_with_nosql_mail:
        if len(_) > 1:
            class_that_update_hmail.create_mails_and_boxes(_, cur, sql_server, con)

    # close connection
    cur.close()


if __name__ == '__main__':
    # all settings here:
    # ldap on default port, if you have another - rewrite HmailMailSync.get_ldap_user connection
    class_that_update_hmail = HmailMailSync(ldap_server='172.16.0.215', ldap_user='vmail',
                                            ldap_password='qwerty-0', search_base='DC=geo,DC=kr')

    # Its pretty simple script, you can edit it as you want
    manual_update(class_that_update_hmail, sql_server='asd.geo.kr', sql_user='sa',
                  sql_database_mail='mail', sql_password='@kizif0eG')




