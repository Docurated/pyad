from adbase import *
import pyadutils

class ADQuery(ADBase):
    # Requests secure authentication. When this flag is set,
    # Active Directory will use Kerberos, and possibly NTLM,
    # to authenticate the client.
    ADS_SECURE_AUTHENTICATION = 1
    # Requires ADSI to use encryption for data
    # exchange over the network.
    ADS_USE_ENCRYPTION = 2
    
    # ADS_SCOPEENUM enumeration. Documented at http://goo.gl/83G1S
    
    # Searches the whole subtree, including all the
    # children and the base object itself.
    ADS_SCOPE_SUBTREE = 2
    # Searches one level of the immediate children,
    # excluding the base object.
    ADS_SCOPE_ONELEVEL = 1
    # Limits the search to the base object.
    # The result contains, at most, one object.
    ADS_SCOPE_BASE = 0

    # ADS_CHASE_REFERRALS_ENUM enumeration. http://msdn.microsoft.com/en-us/library/aa772250(v=vs.85).aspx
    ADS_CHASE_REFERRALS_NEVER        = 0x00
    ADS_CHASE_REFERRALS_SUBORDINATE  = 0x20
    ADS_CHASE_REFERRALS_EXTERNAL     = 0x40,
    ADS_CHASE_REFERRALS_ALWAYS       = 0x20 | 0x40

    # the methodology for performing a command with credentials
    # and for forcing encryption can be found at http://goo.gl/GGCK5
    
    def __init__(self, options={}):
        self.__adodb_conn = win32com.client.Dispatch("ADODB.Connection")
        self.__adodb_conn.Open("Provider=ADSDSOObject")
        if self.default_username and self.default_password:
            self.__adodb_conn.Properties("Encrypt Password").Value = True
            self.__adodb_conn.Properties("User ID").Value = self.default_username
            self.__adodb_conn.Properties("Password").Value = self.default_password
            adsi_flag = ADQuery.ADS_SECURE_AUTHENTICATION | \
                            ADQuery.ADS_USE_ENCRYPTION
            self.__adodb_conn.Properties("ADSI Flag").Value = adsi_flag
            
        self.reset()
    
    def reset(self):
        self.query = None
        self.__rs = self.__rc = None
        self.__queried = False

    def execute_query(self, attributes=["distinguishedName"], where_clause=None,
                    type="LDAP", base_dn=None, page_size=1000, extra_command_properties={}):
        assert type in ("LDAP", "GC")
        if not base_dn:
            if type == "LDAP": 
                base_dn = self._safe_default_domain
            if type == "GC": 
                base_dn = self._safe_default_forest
        self.query = "SELECT %s FROM '%s'" % (','.join(attributes),
                pyadutils.generate_ads_path(base_dn, type,
                        self.default_ldap_server, self.default_ldap_port))
        if where_clause:
            self.query = ' '.join((self.query, 'WHERE', where_clause))
        
        command = win32com.client.Dispatch("ADODB.Command")
        command.ActiveConnection = self.__adodb_conn
        command.Properties("Page Size").value = page_size
        command.Properties("Searchscope").value = ADQuery.ADS_SCOPE_SUBTREE

        for prop, value in extra_command_properties.iteritems():
            command.Properties(prop).value = value
        
        command.CommandText = self.query
        self.__rs, self.__rc = command.Execute()
        self.__queried = True

    def execute_query_range(self, attributes=["distinguishedName"], where_clause=None, base_dn=None, search_scope="subtree"):
        assert type in ("LDAP", "GC")
        if not base_dn:
            base_dn = self._safe_default_domain

        command = win32com.client.Dispatch("ADODB.Command")
        command.ActiveConnection = self.__adodb_conn
        command.Properties("Page Size").Value = page_size

        range_step = 1000
        range_low = 0
        range_high = range_low + range_step - 1
        last_query = False
        end_loop = False
        while not end_loop:
            if last_query:
                command_text = "<LDAP://{0}>;{1};{2};range={3}-*;{5}" % base_dn, where_clause, attributes, range_low, search_scope
                end_loop = True
            else:
                command_text = "<LDAP://{0}>;{1};{2};range={3}-{4};{5}" % base_dn, where_clause, attributes, range_low, range_high, search_scope

            command.CommandText = command_text
            rs, rc = command.Execute()

            if not rs.EOF and rs.Fields[0] is None:
                last_query = True
            else:
                while not rs.EOF
                    d = {}
                    for f in rs.Fields:
                        d[f.Name] = f.Value
                    yield d
                    rs.MoveNext()

            range_low = range_high + 1
            range_high = range_low + range_step - 1

    def get_row_count(self):
        return self.__rs.RecordCount

    def get_single_result(self):
        if self.get_row_count() != 1:
            raise invalidResults(self.get_row_count())
        self.__rs.MoveFirst()
        d = {}
        for f in self.__rs.Fields:
            d[f.Name] = f.Value
        return d

    def get_results(self):
        if not self.__queried:
            raise noExecutedQuery
        if not self.__rs.EOF:
            self.__rs.MoveFirst()
        while not self.__rs.EOF:
            d = {}
            for f in self.__rs.Fields:
                d[f.Name] = f.Value
            yield d
            self.__rs.MoveNext()

    def get_all_results(self):
        if not self.__queried:
            raise noExecutedQuery
        l = []
        for d in self.get_results():
            l.append(d)
        return l
