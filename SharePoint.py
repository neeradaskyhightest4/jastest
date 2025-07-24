from __future__ import absolute_import
from builtins import str
from builtins import range
from skybot.OF.lib.Utilities import dlp_requests as requests
import xml.etree.ElementTree as ET
import time
import os
import re
from skybot.lib.logger import logger
from skybot.OF.lib.core.SkyHighDashboard import Interface
from robot.api.deco import keyword
import json
import random
import string
#from OneDrive import OneDrive
from skybot.OF.lib.core.Services.OneDrive import OneDrive
from robot.libraries.BuiltIn import BuiltIn
from skybot.OF.lib.Utilities.HealthMonitor import trackme
from skybot.OF.lib.Utilities import Utils

from skybot.AM.resources.locators import O365_locators
from skybot.lib.web_automation.CommonHelper import CommonHelper
from skybot.lib.web_automation.ActionsHelper import ActionsHelper
from skybot.lib.web_automation.LocatorType import LocatorType
from skybot.lib.web_automation.SyncHelper import SyncHelper

requests.packages.urllib3.disable_warnings()
retry = 3

class SharePoint(OneDrive):
    file_to_upload = {}
    URL_PATTERN_TO_FIND = 'sharepoint'

    # Folders method

    def get_folder_info(self, folder_id=None, params=""):
        """
        Get info about the folder

        Args:
            folder_id: Id of the folder to be deleted

        Returns:
            List containing JSON of file properties

        Raises:
            None
        """
        if not folder_id:
            folder_id = self.mostrecentfolder
        if folder_id:
            endpoint_GetFolderByServerRelativeUrl = re.sub('GetFolderByServerRelativeUrl(.*)',
                                                                'GetFolderByServerRelativeUrl(\'' + folder_id + '\')',
                                                                    self.endpoint_GetFolderByServerRelativeUrl)
        else:
            endpoint_GetFolderByServerRelativeUrl = self.endpoint_GetFolderByServerRelativeUrl

        if self.site:
            endpoint_GetFolderByServerRelativeUrl = re.sub('(.*)\/_api', self.domain_url + '/sites/' + self.site + '/_api',\
                                                           endpoint_GetFolderByServerRelativeUrl)

        logger.debug("Inside get_folder_info with folder_id: " + str(folder_id))
        get_folder_url = endpoint_GetFolderByServerRelativeUrl + params
        for attempt in range(1, 3):
            headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}
            headers.update({"Accept": "application/json"})
            response_folder_info = requests.get(url=get_folder_url, headers=headers, cookies=self.cookies)
            if response_folder_info.status_code == 200:
                logger.debug(response_folder_info.json())
                return response_folder_info.json()
            elif response_folder_info.status_code in [401, 403]:
                logger.debug("Got a 403 error..")
                self._refresh_token()
                if self.site:
                    self.request_digest = self.req_digest(self.domain_url + "/sites/" + self.site)
            else:
                logger.warn("Failed to get folder info, will retry after 10 sec: " + str(attempt) + "/3")
                time.sleep(10)
                continue
        logger.error("Failed to get folder info post attempt 3 times: " + str(folder_id))
        raise Exception

# Link methods

    @keyword("generate link in ${SERVICE} for last uploaded ${object}")
    def create_link(self, object=object, object_id=None, password="", expiration="", direct="", link_type="edit"):
        """
        Create the link for an object. yet to be implemented

        Args:
            object: File or Folder
            object_id: Object id for which link to be created
            password (Optional): password to be set
            expiration (Optional): expiration to be set
            direct ( Optional): Boolean specifies it is a direct link or not
            link_type ( Optional): edit or view  # specific to sharepoint or onedrive
        Returns:
            Link Id or link that gets generated

        Raises:
            None
        """
        link_url = None
        if object_id is None:
            if object == "file":
                object_id = self.lastuploadedfiles[-1]["fileid"]
            elif object == "folder":
                object_id = self.mostrecentfolder
        logger.debug("Inside create Link to create an anonymous Link for object : " + str(object_id))
        # headers = self.headers.copy()
        headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}
        headers.update({"Content-Type": "application/json"})
        endpoint_create_link = self.domain_url + "/_api/SP.Web.CreateAnonymousLink"
        data = {"url": self.domain_url + object_id, "isEditLink": True}
        for attempt in range(1, 3):
            response_create_link = requests.post(url=endpoint_create_link, headers=headers, data=json.dumps(data), cookies=self.cookies)
            if response_create_link.status_code == 200:
                logger.debug("Response from generating Link is " + response_create_link.text)
                break
            elif response_create_link.status_code in [403, 401]:
                logger.debug("Got a 403 error..")
                self._refresh_token()
                raise Exception
            else:
                logger.debug("Failed to create link, will retry after 10 sec: " + str(attempt) + "/3")
                time.sleep(10)
                continue
        if "value" in response_create_link.json():
            link_url = response_create_link.json()["value"]
            logger.debug("Link is generated with the link_url: " + link_url)
        return link_url

    def _get_endpoints(self):
        """
        Going to build endpoints based on the email/domain
        :return: None
        """
        default_root_folder = "/Shared Documents"
        try:
            for each in self.shutil.services.get("allservices"):
                if each == "SharePoint":
                    self.root_folder = self.shutil.services.get("allservices").get("SharePoint")\
                        .get("library", default_root_folder)
                    break
                else:
                    if "_" in each:
                        if each.split("_")[1] == str(self.instance_id):
                            self.root_folder = self.shutil.services.get("allservices").get(each) \
                                .get("library", default_root_folder)
                            break
        except Exception as e:
            logger.warn("Not able to get default library due to: ${0}, using hardcoded default ${1}".format(e, str(default_root_folder)))
            self.root_folder = default_root_folder
        self.user_flat = self.user.replace("@", ".").replace(".", "_")
        self.domain_url = "https://" + self.domain_name + ".sharepoint.com"
        self.domain_admin_url = "https://" + self.domain_name + "-admin.sharepoint.com"
        self.endpoint_GetFolderByServerRelativeUrl = self.domain_url + "/_api/Web/GetFolderByServerRelativeUrl(\'" + \
                                                        self.root_folder + "\')"
        self.endpoint_GetFileByServerRelativeUrl = self.domain_url + "/sites/TestAuto/_api/Web/GetFileByServerRelativeUrl(\'" + \
                                                    self.root_folder + "\')"
        self.endpoint_users = "https://graph.microsoft.com/v1.0/" + self.domain_name + ".onmicrosoft.com" + "/users"
        self.endpoint_retrieve_links = "https://" + self.domain_name + ".sharepoint.com" + \
                                       "/_vti_bin/client.svc/ProcessQuery"
        self.endpoint_contextinfo = self.domain_url + '/sites/QAAutomationPublic/_api/contextinfo'
        self.endpoint_groups = "https://graph.microsoft.com/v1.0/" + self.domain_name + ".onmicrosoft.com" + "/groups"
        self.endpoint_create_field = self.domain_url + "/_api/web/lists/getbytitle(\'Documents\')/Fields"
        self.endpoint_create_link = self.domain_url + "/_api/"
        logger.debug("Sharepoint Library documents been used is: " + str(self.endpoint_GetFolderByServerRelativeUrl))
        self.endpoint_retrieve_flexilink = self.domain_url + \
            "/_api/web/getlistitem(@url)/getsharinginformation/permissionsInformation/links?@url='%s'"
        self.endpoint_host_web_url = self.domain_url
        self.endpoint_list = self.domain_url + "/_api/SP.AppContextSite(@target)/web/Lists"
        self.endpoint_create_list = self.endpoint_list + "?@target='" + self.domain_url + "'"
        self.endpoint_GetFileByServerRelativePath = self.domain_url + \
                                "/_api/web/GetFileByServerRelativePath(decodedurl=@relativeUrl)/$value?@relativeUrl='"
        #self.default_root_folder = default_root_folder
        self.default_root_folder = self.root_folder
        self.endpoint_DirectAccessSharing = self.domain_url + "/_api/SP.Sharing.DocumentSharingManager.UpdateDocumentSharingInfo"
        self.endpoint_sharepoint_group = self.domain_url + '/sites/QAAutomationPublic/_api/web/sitegroups'
        self.endpoint_SPGroup_add_users = self.domain_url + "/sites/QAAutomationPublic/_api/web/sitegroups/GetById({0})/users"
        #self.endpoint_SPGroup_add_users = self.domain_url + "/sites/QAAutomationPublic/_api/web/sitegroups/GetById({0})/users"
        self.endpoint_UserCreated_SPGroup_add_users = self.domain_url + "/sites/{1}/_api/web/sitegroups/GetById({0})/users"
        self.endpoint_SPGroup_get_users = self.domain_url + "/_api/web/sitegroups/GetByName('" + "{0}" + "')"
        self.endpoint_SPGroup_delete = self.domain_url + "/_api/web/sitegroups/removebyid({0})"
        self.endpoint_groups_url = "https://graph.microsoft.com/v1.0/groups/"
        self.endpoint_Flexilink = self.domain_url + "/_api/web/GetListItemUsingPath(decodedurl=@u)/ShareLink?@u='{0}'"
        self.endpoint_Folder_DirectAccessSharing = self.domain_url + \
                    "/sites/QAAutomationPublic/_api/web/GetFolderByServerRelativeUrl(@relativeUrl)/ListItemAllFields/ShareObject?@relativeUrl='%s'"
        self.endpoint_File_DirectAccessSharing = self.domain_url + \
                    "/sites/QAAutomationPublic/_api/web/GetFileByServerRelativeUrl(@relativeUrl)/ListItemAllFields/ShareObject?@relativeUrl='%s'"
        self.endpoint_DirectAccessSharing_listId = self.domain_url + "/sites/QAAutomationPublic/_api/web/Lists(@a1)/GetItemById(@a2)/ShareObject?@a1='{%s}'&@a2='%s'"
        self.endpoint_GetFileListByServerRelativePathUrl = self.domain_url + "/_api/web/GetFileByS  erverRelativePath(decodedurl=@relativeUrl)" + \
                                                           "/ListItemAllFields?@relativeUrl='%s'"
        self.endpoint_GetFolderListByServerRelativePathUrl = self.domain_url + "/_api/web/GetFolderByServerRelativePath(decodedurl=@relativeUrl)" + \
                                                             "/ListItemAllFields?@relativeUrl='%s'"
        self.endpoint_Flexilink_bylistid = self.domain_url + "/_api/web/Lists(@a1)/GetItemById(@a2)/ShareLink?@a1='{%s}'&@a2='%s'"

    def _get_SPgroups_id(self,groupname):
        headers = {"Authorization": "Bearer " + self.access_token, "content-type": "application/json;odata=verbose","Accept": "application/json;odata=verbose"}
        result = requests.get(url=self.endpoint_sharepoint_group, headers=headers)
        if result.status_code in [200,201]:
            for group in json.loads(result.text)["d"]["results"]:
                if group["Title"]==groupname:
                    return group["Id"]
            logger.error("Group %s not Found!" %(groupname))
            return False
        elif result.status_code in [401,403]:
            logger.debug("Retrying after refreshing access token")
            self._refresh_token
            r = self._get_SPgroups_id(groupname)
            return r
        else:
            ("Unable to Fetch all groups due to " + str(result.text))
            return False

    @keyword("In ${SERVICE} get external members from ${members_list}")
    def get_external_users_from_group(self,members_list):
        external_users=[]
        for i in members_list:
            user_domain=i.split("@")[1].split(".")[0]
            logger.debug(user_domain)
            logger.debug(self.domain_name)
            if user_domain==self.domain_name:
                continue
            else:
                external_users.append(i)
        return external_users


    @keyword("In ${SERVICE} create sharepoint group ${group_name}")
    def create_sharepoint_group(self,group_name):

        headers = {"Authorization": "Bearer " + self.access_token, "Accept": "application/json;odata=verbose","content-type": "application/json;odata=verbose"}
        data={
             "__metadata": {
            "type": "SP.Group"
            },
            "Title": group_name,
            "Description": "Automation"
        }
        result=requests.post(url=self.endpoint_sharepoint_group, headers=headers, data=json.dumps(data))
        if result.status_code == 201:
            json_result = (json.loads(result.text))
            GroupId=json_result["d"]["Id"]
            self.lastcreatedgroup.append(
                {
                    "groupid": str(GroupId),
                    "groupname": str(group_name),
                    "email":str(group_name)
                }
            )
            return True
        elif result.status_code == 500:
            logger.debug("Group %s already exists, deleting and try again " %(group_name))
            deleted = self.delete_sharepoint_group(group_name)
            if deleted:
                r = self.create_sharepoint_group(group_name)
                return r
            else:
               return False
        elif result.status_code in [401,403]:
            logger.debug("Retrying after refreshing access token")
            self._refresh_token()
            r = self.create_sharepoint_group(group_name)
            return r
        else:
            logger.error("Unable to create Sharepoint group due to " + str(result.text))
            return False

    @keyword("In ${SERVICE} validate for external users {usersList}")
    def validate_for_external_users(self,usersList):
        logger.debug(usersList)
        domains_list=[]
        user_domain = self.domain_url.split("//")[1]
        for user in usersList:
            domain=user.split('@')[1].split(".")[0]
            domains_list.append()
            domain==user_domain


    @keyword("In ${SERVICE} get members of sharepoint group ${groupname}")
    def get_members_of_sharepoint_group(self,groupname=None):
        for group in self.lastcreatedgroup:
            if group['groupname'] == groupname:
                groupid = group["groupid"]
                break
        url = self.endpoint_SPGroup_get_users.format(groupname) + "/users?$select=Email,Id"
        logger.debug(url)
        headers = {"Authorization": "Bearer " + self.access_token, "Accept": "application/json;odata=nometadata","content-type": "application/json;odata=nometadata"}
        result = requests.get(url=url, headers=headers)
        if result.status_code in [200, 201]:
            my_json = result.content.decode('utf8')
            data = json.loads(my_json)
            users = []
            results = data['value']
            for i in results:
                users.append(i['Email'])
            return users
        elif result.status_code in [401, 403]:
            logger.debug("Retrying after refreshing access token")
            self._refresh_token
            result = requests.get(url=url, headers=headers)
            return result.content
        else:
            ("Unable to Fetch members of group due to " + str(result.text))
            return False

    @keyword("In ${SERVICE} add members to sharepoint site ${sitename} default group ${groupid}")
    def add_members_to_sharepoint_site_default_group(self,sitename=None, groupid=None ):
        add_result = []
        users = BuiltIn().replace_variables('${o365_users}')
        self.members_to_collaborate = self.get_users_for_O365Group(users)
        if sitename == "":
            url = self.endpoint_SPGroup_add_users.format(groupid)
        else:
            url = self.endpoint_UserCreated_SPGroup_add_users.format(groupid,sitename)
        headers = {"Authorization": "Bearer " + self.access_token, "content-type": "application/json;odata=verbose"}
        for user in self.members_to_collaborate:
            login_name = self._get_loginname(user)
            data = {
                    '__metadata': {

                    'type': 'SP.User'
                },
                'LoginName': login_name
            }

            response=requests.post(url=url, headers=headers, data=json.dumps(data))
            if response.status_code == 201:
                add_result.append({user:True})
            elif response.status_code in [401,400]:
                self._refresh_token()
                response = requests.post(url=url, headers=headers, data=json.dumps(data))
                if response.status_code == 201:
                    add_result.append({user: True})
            else:
                add_result.append({user: False})
                logger.debug(add_result)
                logger.error("Adding member %s failed " %(user))
                return False

        return True

    @keyword("In ${SERVICE} add members to sharepoint group ${groupname}")
    def add_members_to_sharepoint_group(self,groupname=None ):
        add_result = []
        users = BuiltIn().replace_variables('${o365_users}')
        self.members_to_collaborate = self.get_users_for_O365Group(users)
        index=0
        if groupname:
            for group in self.lastcreatedgroup:
                if group['groupname']== groupname:
                    groupid=group["groupid"]
                    self.lastcreatedgroup[index]['members'] = self.members_to_collaborate
                index +=1
        else:
            groupid = self.lastcreatedgroup[-1].get("groupid")
        if not groupid:
            logger.error("Group name %s is not created" % (groupname))
            return False
        url = self.endpoint_SPGroup_add_users.format(groupid)
        headers = {"Authorization": "Bearer " + self.access_token, "content-type": "application/json;odata=verbose"}
        for user in self.members_to_collaborate:
            login_name = self._get_loginname(user)
            data = {
                    '__metadata': {

                    'type': 'SP.User'
                },
                'LoginName': login_name
            }

            response=requests.post(url=url, headers=headers, data=json.dumps(data))
            if response.status_code == 201:
                add_result.append({user:True})
            elif response.status_code in [401,400]:
                self._refresh_token()
                response = requests.post(url=url, headers=headers, data=json.dumps(data))
                if response.status_code == 201:
                    add_result.append({user: True})
            else:
                add_result.append({user: False})
                logger.debug(add_result)
                logger.error("Adding member %s failed " %(user))
                return False

        return True

    @keyword("In ${SERVICE} Get groupID from ${groupname}")
    def get_groupid(self,groupname=None):
        groupid = None
        groupid = self._get_SPgroups_id(groupname)
        return groupid


    @keyword("In ${SERVICE} delete latest sharepoint group")
    def delete_sharepoint_group(self,groupname=None):
        headers = {"Authorization": "Bearer " + self.access_token, "content-type": "application/json;odata=verbose"}
        delete_result = []
        groupid = None
        retry = 0
        if groupname:
            for each in self.lastcreatedgroup:
                if each['groupname'] == groupname:
                    groupid = each.get("groupid")
                    break

            if not groupid:
                groupid = self._get_SPgroups_id(groupname)

            if groupid:
                url = self.endpoint_SPGroup_delete.format(groupid)
                response = requests.post(url=url, headers=headers)
                if response.status_code in [200, 201]:
                    delete_result.append(True)
                elif response.status_code in [401,403] and retry == 0:
                    self._refresh_token()
                    response = requests.post(url=url, headers=headers)
                    if response.status_code in [200, 201]:
                        delete_result.append(True)
                    else:
                        delete_result.append(False)
        else:
            for group in self.lastcreatedgroup:
                retry = 0
                groupid = group.get("groupid")
                url = self.endpoint_SPGroup_delete.format(groupid)
                result = requests.post(url=url, headers=headers)
                if result.status_code in [200, 201]:
                    delete_result.append(True)
                elif result.status_code in [401, 403] and retry == 0:
                    self._refresh_token()
                    result = requests.post(url=url, headers=headers)
                    if result.status_code in [200, 201]:
                        delete_result.append(True)
                    else:
                        delete_result.append(False)
                else:
                    logger.error("Unable to delete SP Group %s due to %s and status %s " % (
                            groupid, result.text, result.status_code))
                    delete_result.append(False)

        return all(delete_result)

    def _get_loginname(self,username):
        if username:
            for user in self.response_get_all_users.json().get("value"):
                if not user.get("mail"):
                    continue
                if username == user.get("mail"):
                    loginname="i:0#.f|membership|"+str(user.get("userPrincipalName").lower())
                    return loginname
            logger.debug("user %s not found in Sharepoint" %(username))
            return False
        else:
            logger.debug("username sent is None")
            return False



    def upload_file(self, filename, parent_id=0, overwrite=True, site=None):

        headers = {"X-RequestDigest": self.request_digest, "Content-Type": self.get_mime_type(filename=os.path.basename(filename)),
                    "Accept": "application/json"}

        filename = str(filename)
        if self.testdata not in filename:
            filename = self.testdata + "/" + filename
        file = str(os.path.basename(filename))

        with open(filename, "rb") as fp:
            if filename not in SharePoint.file_to_upload:
                SharePoint.file_to_upload[file] = fp.read()
        parent_id, self.mostrecentfolder = [0 if self.mostrecentfolder is None else self.mostrecentfolder] * 2

        if parent_id:
            endpoint_GetFolderByServerRelativeUrl = re.sub('GetFolderByServerRelativeUrl(.*)',
                                                           'GetFolderByServerRelativeUrl(\'' + parent_id + '\')',
                                                           self.endpoint_GetFolderByServerRelativeUrl)
        else:
            endpoint_GetFolderByServerRelativeUrl = self.endpoint_GetFolderByServerRelativeUrl
        upload_url = endpoint_GetFolderByServerRelativeUrl + '/Files/add(url=\'' + requests.encode_url(os.path.basename(filename)) \
                            + '\', overwrite=true)'
        if site:
            self.site = site[type(self).__name__]
            req_dig = self.req_digest(self.domain_url + "/sites/" + self.site)
            headers["X-RequestDigest"] = req_dig
            upload_url = re.sub('(.*)\/_api', self.domain_url + '/sites/' + self.site + '/_api', upload_url)
        else:
            self.site = None

        logger.info("Create File url is " + upload_url)

        for i in range(retry):
            response = requests.post(upload_url, headers=headers,cookies=self.cookies, data=SharePoint.file_to_upload[file])
            self.response = response
            if response.status_code == 200:
                break
            if response.status_code in [401,403]:
                logger.debug("Got a 403 error refreshing access token...")
                self._refresh_token()
                if site:
                    req_dig = self.req_digest(self.domain_url + "/sites/" + self.site)
                    headers["X-RequestDigest"] = req_dig

        if response.status_code != 200:
            raise Exception

        logger.info("Response post upload file is: " + response.text)
        if "ServerRelativeUrl" in response.json():
            logger.info("File " + os.path.basename(filename) + " is successfully uploaded")
            file_id = response.json()["ServerRelativeUrl"]

        if self.site:
            quarantineref = "/sites/" + self.site + ":" + str(file_id)
        else:
            quarantineref = "/:" + str(file_id)

        self.lastuploadedfiles.append(
            {
                "fileid":str(file_id),
                "filename":str(file),
                "folderid": parent_id,
                "quarantineref": quarantineref,
                "permissions_object": {"id": self.mostrecentfolder, "permissions_list": None}

            }
        )
        logger.info("Files uploaded thus far: " + str(self.lastuploadedfiles))
        test_name = BuiltIn().replace_variables('${TEST_NAME}')
        BuiltIn().set_suite_metadata(test_name + "_" + str(self.instance_id) + "_lastuploadedfiles", self.lastuploadedfiles)
        return file_id

    def req_digest(self, site):
        url = site + '/_api/contextinfo'
        headers = {"Accept": "application/json", "Content-Length": '0'}
        response = requests.post(url, headers=headers, cookies=self.cookies)
        return response.json()['FormDigestValue']


    @keyword("get different user from ${SERVICE}")
    def get_different_user_for_collab(self):
        users = self.get_all_users()
        users.remove(str(self.user))
        return random.choice(users)

    def enable_api_access(self,params,driver_obj,api, wait, EC, By):
        time.sleep(5)
        wait.until(EC.visibility_of_element_located((By.XPATH, api.page_elements_dict["common"]["preReqCheck"])))
        driver_obj.find_element_by_xpath(api.page_elements_dict['common']['preReqCheck']).click()
        logger.debug("Clicked Prerequisites")
        time.sleep(3)
        wait.until(EC.element_to_be_clickable((By.XPATH, api.page_elements_dict["common"]["nextButton"])))
        driver_obj.find_element_by_xpath(api.page_elements_dict['common']['nextButton']).click()
        logger.debug("Clicked Next")
        time.sleep(5)
        driver_obj.find_element_by_xpath(api.page_elements_dict['common']['credsButton']).click()
        logger.debug("Clicked Provide Credentials")
        time.sleep(5)
        handles=driver_obj.window_handles
        current=driver_obj.current_window_handle
        driver_obj.switch_to.window(handles[1])
        driver_obj.find_element_by_xpath(api.page_elements_dict['SharePoint']['adminResourceURL']).send_keys(
            str(params['resourceURL']))
        driver_obj.find_element_by_xpath(api.page_elements_dict['Jive']['JiveSubmit']).click()
        time.sleep(5)
        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['oneDriveEmail']).send_keys(
            str(params['email']))
        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['nextButton']).click()
        time.sleep(10)
        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['password']).send_keys(
            str(params['password']))
        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['nextButton']).click()
        time.sleep(10)
        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['acceptButton']).click()
        time.sleep(10)
        driver_obj.switch_to.window(current)
        return True

    @keyword("In ${SERVICE} create o365 group with ${visibility} ${name}")
    def create_o365_group(self,visibility,name,type=None):
        result = True
        #name = name+".sharepoint"
        if not type:
            type = ["Unified"]
            logger.debug("type of group is === " + str(type))
        else:
            logger.debug("type of group in else is === " + str(type))
        existing_groups = self.get_o365_groups()
        logger.debug("existing groups === " + str(existing_groups))

        for group in existing_groups["value"]:
            logger.debug("group is ==" + str(group))
            if group.get("displayName") == name:
                logger.debug("O365 groups named %s already exists " % (name))
                self.lastcreatedgroup.append(
                                        {
                                            "groupid":str(group.get("id")),
                                            "groupname":str(group.get("displayName")),
                                            "email":str(group.get("mail")),
                                            "id": str(group.get("id")),
                                            'fid': "c:0o.c|federateddirectoryclaimprovider|" + str(group.get("id")),
                                            'apiDisplayText': str(group.get("displayName")) + " Members",
                                            'apiDisplayName': str(group.get("displayName")) + " Members"

                                        }
                                     )
                logger.debug("last created group when group already exists == " + str(self.lastcreatedgroup))
                return result

        owner_id= self._get_user_ids([self.admin])

        logger.debug("===Group not found, creating new one===")
        mail_nick_name=name.replace(" ","")
        headers = {"Authorization": "Bearer " + self.access_token_graph, "content-type": "application/json"}
        data={
              "groupTypes": type,
              "displayName": name,
              "mailNickname": mail_nick_name,
              "mailEnabled": "true",
              "securityEnabled": "false",
              "visibility":visibility,
              "owners@odata.bind":["https://graph.microsoft.com/v1.0/users/"+ owner_id[0]]
              }
        logger.debug("endpoint is %s, headers are %s. data is %s " % (self.endpoint_groups,headers,data))
        #logger.debug("Data type is" + str(type(data)) )
        #logger.debug("headers type is" + str(type(headers)) )
        response_create_group = requests.post(url=self.endpoint_groups, headers=headers, data=json.dumps(data))
        new_group=response_create_group.json()
        logger.debug("response is " + str(new_group))
        if response_create_group.status_code in (200,201):

            self.lastcreatedgroup.append(
                                            {
                                                "groupid":str(new_group.get("id")),
                                                "groupname":str(new_group.get("displayName")),
                                                "email":str(new_group.get("mail")),
                                                "id":str(new_group.get("id")),
                                                'fid': "c:0o.c|federateddirectoryclaimprovider|"+ str(new_group.get("id")),
                                                'apiDisplayText': str(new_group.get("displayName")) + " Members",
                                                'apiDisplayName': str(new_group.get("displayName")) + " Members"
                                            }
                                         )
            logger.debug("last created group when new group is created  == " + str(self.lastcreatedgroup))
        else:
            logger.error("Group creation failed due to " + str(response_create_group._content))
            result=False
        return result
    @keyword("In ${SERVICE} verify retry count in case of rate limits")
    def verify_retrycount_for_ratelimits(self):
        import redis
        redis_server = self.shutil.get_redis_ip()
        redis_client = redis.Redis(redis_server)
        redis_key = 'offlinedlp:event_error_metrics:{0}:{1}:{2}:event_retry_count'.format(self.tenantid, self.cspid, self.instance_id)
        file_path = self.lastuploadedfiles[-1]['quarantineref']
        retries = int(redis_client.hget(redis_key, file_path))
        if retries >= 2:
            logger.debug("Retrying in case of rate limits")
            return True
        else:
            logger.debug("Not retrying in case of rate limits")
            return False

    @keyword("Log into ${service} as ${user}")
    def login_to_service_ui(self, user):
        logger.console("Logging into SharePoint UI now")
        if super(SharePoint, self).login_to_service_ui(user):
            logger.console("Successfully Logged into SharePoint!")
            return True
        else:
            logger.console("Could not login to SharePoint")
            return False


    def click_url(self, url):
        url = [x for sublist in url for x in sublist]
        url = list(dict.fromkeys(url))
        urls= [value for value in url if value is not False]

        try:
            link_to_click = self.select_url(urls, SharePoint.URL_PATTERN_TO_FIND)
        except Exception as e:
            logger.console("Link received is None")
            return False

        for link in link_to_click:
            if ('-my.' in link)== False and '.sharepoint' in link:
                link_to_click=link
        if link_to_click:
            logger.console("Link to click in SharePoint= " + str(link_to_click))
            CommonHelper.go_to_url(self.driver, link_to_click)
            #waiting for the page to load
            CommonHelper.wait_for_seconds(8)
            if CommonHelper.is_element_displayed(self.driver, LocatorType.XPATH,
                                                 O365_locators['onedrive_item_removed_page']):
                logger.console("SharePoint link is expired! The user does not have permission to access this file")
                return True
            else:
                logger.console("SharePoint link accessible")
                return False
        else:
            logger.error("No SharePoint link received")
            return False


if __name__ == '__main__':
    from skybot.lib import SHNInterface
    SHNInterface.myenv = SHNInterface.Util("qaautoregression", "dlpqap1@gmail.com", "Welcome2dlp#")
    from skybot.OF.lib.core.SkyHighDashboard.ShnDlpInterface import ShnDlpUtil
    shutil = ShnDlpUtil("qaautoregression", 5642, "Welcome2dlp#", "dlpqap1@gmail.com", None, None, use_token=True)
    SHNInterface.myenv = shutil
    shutil.current_service = "SharePoint"
    os.environ.setdefault("office365_password", "")

    Od = SharePoint(shutil, "qaautoregression", 5642,16131, "admin@shnqaeu4.onmicrosoft.com", instance_id=12997)

    Od.as_user("user1@shnqaeu4.onmicrosoft.com")
    #Od.create_folder('Test7',site='automationgroupsite')
    #Od.upload_file('Confidential.docx',site='automationgroupsite')
    #Od.create_folder('FoldertoCheckCollaboration3')
    #Od.upload_file('/Users/siddharth/Documents/DlpProjectOF/trunk/DLPRobotFramework/data/files/forbidden.txt')
    #Od.add_permission(None,user_attr={'role':'editor','email':'*'},file_collaboration=False)
    # Od.as_user("admin@ak001.onmicrosoft.com")
    # Od.create_folder("myfolder", "/personal/admin_ak001_onmicrosoft_com/Documents")
    # Od.create_folder("Ashish1_Folder")
    # Od.upload_file("/Users/ashishk/Documents/Skyhigh/svn/automation/automation/DLPFramework_Current/DLPRobotFramework/data/files/test.txt")
    # Od.list_permissions("/personal/admin_ak5_onmicrosoft_com/Documents/My Folder")
    object_id = "/Shared Documents/1596647468.433655/Confidential.docx"
    user_attr = {"email": "*", "role": "editor"}
    Od.add_permission(object_id, user_attr,add_to_all_collaborators=False)
    # # # Od.get_drive()
    # print Od.upload_file("/Users/ashishk/Documents/Skyhigh/svn/automation/automation/DLPFramework_Current/DLPRobotFramework/data/files/test.txt")
    # Od.delete_file(file_id="my folder/test.txt")
