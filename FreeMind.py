# -*- coding: utf-8 -*- 

import argparse
import logging.config
import sys
import os
from xml.etree import ElementTree as ET
import xml.dom.minidom as minidom
import xml.etree.cElementTree  as lxmlET

import testlink
from xlrd import open_workbook
import pprint

PKG_PATH = './'

TC_ID = 0
TC_TITLE = 1
TC_REQ_LINKS = 2
REQ_LINK_SPEC = 0
REQ_LINK_ID = 1
REQ_LINK_TITLE = 2

REQ_ID = 0
REQ_TITLE = 1
REQ_DESC = 2
REQ_VER_TEAM = 3
REQ_COMMENT = 4
REQ_PHASE = 5

PREFIX_TITLE_SEP = '::'

''' The following functions (CDATA and _serialize_xml) is a workaround for using CDATA section with ElementTree
'''


def CDATA(text=None):
    element = ET.Element('![CDATA[')
    element.text = text
    return element


_original_serialize_xml = ET._serialize_xml


def _serialize_xml(write, elem, encoding, qnames, namespaces):
    if elem.tag == '![CDATA[':
        #write("<%s%s]]>%s" % (elem.tag, elem.text, elem.tail))
        write("<%s%s]]>" % (elem.tag, elem.text))
        return
    return _original_serialize_xml(
        write, elem, encoding, qnames, namespaces)


ET._serialize_xml = ET._serialize['xml'] = _serialize_xml


class FreeMind(object):
    ''' This is a class working with TestLink and various offline templates.
        Basically it includes the features of generating TDS, linking TDS with test cases and test plans.
        It also have some advance features including generating PMR, PFS and traceability from Excel template.
        Eventually it has the features of create test plan in test link and export various test reports from TestLink.
        Please check the related configuration file and work instructions for detailed information.
    '''

    def __init__(self, logger, cfg_file=None):
        self.logger = logger
        self.log_prefix = 'FreeMind:'
        self.fm_tree = None
        self.fm_file = None
        self.tc_tree = None
        self.tc_file = None
        self.node_found = False

        self.testlink_url = None
        self.testlink_devkey = None
        self.tls = None
        self.tc_prefix = None
        self.project_name = None
        self.pfs_prefix = None
        self.pmr_prefix = None
        self.tds_prefix = None
        self.test_plan = None
        self.requirements_url = None
        self.pmr_url = None
        self.pfs_url = None
        self.tds_url = None
        self.tp_url = None
        self.tc_url = None
        self.based_tp_url = None
        self.flashobject_swf = None
        self.flashobject_js = None
        self.html_template = None
        self.logger.info(self.log_prefix + \
                         "FreeMind Tool V0.1 for Test Design and Test Management.")
        if cfg_file:
            # Parse the configuration file automatically if it's specified
            self.logger.info(self.log_prefix + \
                             "Parse the configuration file (%s)." % \
                             (cfg_file))
            self._parse_cfg_file(cfg_file)

    def _parse_cfg_file(self, cfg_file):
        cfg_tree = ET.parse(cfg_file)
        cfg_root = cfg_tree.getroot()

        # Firstly get all configurations from the default configuration file
        for item in cfg_root.iter():
            if item.tag == 'testlink':
                self.testlink_rpc_url = item.attrib['URL'].strip()
                self.testlink_url = '/'.join(self.testlink_rpc_url.split('/')[:3])
                self.testlink_devkey = item.attrib['DEV_KEY'].strip()
                os.environ['TESTLINK_API_PYTHON_SERVER_URL'] = self.testlink_rpc_url
                os.environ['TESTLINK_API_PYTHON_DEVKEY'] = self.testlink_devkey
            if item.tag == 'repository':
                self.repo_prefix = item.attrib['PREFIX'].strip()
                self.repo_name = item.attrib['NAME'].strip()
            if item.tag == 'project':
                self.project_name = item.attrib['NAME'].strip()
                self.pfs_prefix = item.attrib['PFS_PREFIX'].strip()
                if self.pfs_prefix == "":
                    self.pfs_prefix = self.project_name + '_'
                self.pmr_prefix = item.attrib['PMR_PREFIX'].strip()
                if self.pmr_prefix == "":
                    self.pmr_prefix = self.project_name + '_'
                self.tds_prefix = item.attrib['TDS_PREFIX'].strip()
                if self.tds_prefix == "":
                    self.tds_prefix = self.project_name + '_'

            if item.tag == 'file_location':
                file_location = item.attrib['URL'].strip()
            if item.tag == 'requirements_url':
                self.requirements_url = self._get_url(file_location, item.text.strip())
            if item.tag == 'pmr_url':
                self.pmr_url = self._get_url(file_location, item.text.strip())
            if item.tag == 'pfs_url':
                self.pfs_url = self._get_url(file_location, item.text.strip())
            if item.tag == 'tds_url':
                self.tds_url = self._get_url(file_location, item.text.strip())
            if item.tag == 'tc_url':
                self.tc_url = self._get_url(file_location, item.text.strip())
            if item.tag == 'tp_url':
                self.tp_url = self._get_url(file_location, item.text.strip())
            if item.tag == 'based_tp_url':
                self.based_tp_url = self._get_url(file_location, item.text.strip())

            if item.tag == 'freemind':
                freemind = item.attrib['URL'].strip()
            if item.tag == 'flashobject_swf':
                self.flashobject_swf = freemind + item.text.strip()
            if item.tag == 'flashobject_js':
                self.flashobject_js = freemind + item.text.strip()
            if item.tag == 'html_template':
                self.html_template = freemind + item.text.strip()

                # Secondly perform all enabled actions.
        for action in cfg_root.iter('action'):
            if action.attrib['ENABLE'].strip() <> '1':
                continue
            action_name = action.attrib['NAME'].strip()
            self.logger.info(self.log_prefix + \
                             "Perform the enabled action (%s) specified in the configuration file (%s)." % \
                             (action_name, cfg_file))
            if action_name == 'Extract_Requirements':
                self.extract_requirements(self.requirements_url, action.attrib['TEMPLATE'].strip())
            if action_name == 'Link_PFS_with_PMR':
                pass  #self.link_pfs_pmr(self.pmr_url, self.pfs_url)
            if action_name == 'Link_PFS_with_TCs':
                pass  #self.link_tc2pfs(self.pfs_url, self.tc_url)
            if action_name == 'Generate_TDS':
                self.gen_tds(self.tds_url, action.attrib['REMOVE_PREFIX'].strip())
            if action_name == 'Link_TDS_with_TCs':
                self.link_tc2tds(self.tds_url, self.tc_url)
            if action_name == 'Link_TDS_with_TCs-TPs':
                self.link_tp2tds_tc(self.tds_url, self.tc_url, action.attrib['FILTER'].strip())
            if action_name == 'Link_TDS_with_TCs-PFS':
                self.link_pfs2tds(self.tds_url, self.tc_url, self.pfs_url)
            if action_name == 'Link_TCs_with_TDS':
                self.link_tds2tc(self.tc_url, self.tds_url)
            if action_name == 'Create_Test_Plan':
                self.create_test_plan(self.tp_url, action.attrib['AUTO'].strip(), action.attrib['TEAM'].strip())

        return 0

    def _get_url(self, file_location, file_name):
        ''' Combine the file location path with file names if they are sharing the same file location.
        '''
        res = None
        if os.path.exists(file_location):
            res = file_location + file_name
        else:
            res = file_name

        return res

    def parse_freemind(self, file_name):
        self.fm_tree = ET.parse(file_name)
        self.fm_file = file_name
        return 0

    def _gen_freemind(self):
        self.fm_tree.write(os.path.splitext(self.fm_file)[0] + "_New.mm")
        return 0

    def add_prefix(self, file_name):
        self.parse_freemind(file_name)
        self._add_node_prefix(self.fm_tree.getroot(), '0')
        self._gen_freemind()
        return 0

    def remove_prefix(self, file_name):
        self.parse_freemind(file_name)
        self._remove_node_prefix(self.fm_tree.getroot())
        self._gen_freemind()
        return 0

    def gen_tds(self, file_name, remove_prefix):
        tds_item_list = ['TDS', []]
        fm_tree = ET.parse(file_name)
        tds_root = fm_tree.getroot()
        #Firstly remove all prefix hence we will number them again.
        self._remove_node_prefix(tds_root)

        self.logger.info(self.log_prefix + \
                         "Read TDS file (%s) and get the information of last nodes which will be used to generate the xml file for importing to TestLink" % \
                         (file_name))
        self._get_tds_items(tds_root, '0', '', tds_item_list[1])

        filename = os.path.splitext(file_name)[0] + '.xml'
        title = os.path.splitext(os.path.split(file_name)[-1])[0]
        self._gen_req_xml([tds_item_list], title, filename, self.tds_prefix)

        if remove_prefix == '1':
            self._remove_node_prefix(tds_root)
            fm_tree.write(file_name)

        return 0

    def _get_tds_items(self, node, num, desc, item_list):
        res = 0
        i = 0
        prefix = ''
        content = ''
        for child in node:
            if child.tag == 'node':
                i = i + 1
                prefix = str(num) + '.' + str(i)
                content = desc + '|' + child.attrib['TEXT']
                # Make sure this is not the test case or requirement link node since only they are nodes with links
                if not child.attrib.has_key('LINK'):
                    # If this is the last node or the node has sub-nodes with links, consider it as a TDS item
                    if (child.find('node') == None) or (
                                (child.find('node') <> None) and (child.find('node').attrib.has_key('LINK'))):
                        item_list.append([prefix[4:], '|'.join(content.split('|')[2:]), '', 'SIT'])
                    self._get_tds_items(child, prefix, content, item_list)

        return res

    def _gen_req_xml(self, item_list, doc_title, filename, prefix, relation_list=None):
        ''' item_list is a list like [GROUP_NAME, [ [REQ_ID, REQ_TITLE, REQ_DESC, REQ_VER_TEAM], ... ] ]
        '''
        res = 0

        self.logger.info(self.log_prefix + \
                         "Generating the xml file %s (Document Title: %s. Document ID Prefix: %s) for importing to TestLink." % \
                         (filename, doc_title, prefix))

        tds = ET.Element('requirement-specification')
        #title = ''.join(self.fm_file.split('.')[:-1])
        req_spec = ET.SubElement(tds, 'req_spec', {'title': doc_title, 'doc_id': doc_title})
        req_type = ET.SubElement(req_spec, 'type')
        req_type.append(CDATA(2))
        node_order = ET.SubElement(req_spec, 'node_order')
        node_order.append(CDATA(1))
        total_req = ET.SubElement(req_spec, 'total_req')
        total_req.append(CDATA(0))
        scope = ET.SubElement(req_spec, 'scope')
        scope.append(CDATA(''))

        #pprint.pprint(item_list)
        i = 0
        for group in item_list:
            for item in group[1]:
                i = i + 1
                requirement = ET.SubElement(req_spec, 'requirement')
                docid = ET.SubElement(requirement, 'docid')
                docid.append(CDATA(prefix + item[REQ_ID]))
                title = ET.SubElement(requirement, 'title')
                title.append(CDATA(item[REQ_TITLE]))
                node_order = ET.SubElement(requirement, 'node_order')
                node_order.append(CDATA(i))
                description = ET.SubElement(requirement, 'description')
                description.append(CDATA(item[REQ_DESC]))
                status = ET.SubElement(requirement, 'status')
                status.append(CDATA('V'))
                req_type = ET.SubElement(requirement, 'type')
                req_type.append(CDATA(2))
                expected_coverage = ET.SubElement(requirement, 'expected_coverage')
                expected_coverage.append(CDATA(1))
                custom_fields = ET.SubElement(requirement, 'custom_fields')
                custom_field = ET.SubElement(custom_fields, 'custom_field')
                name = ET.SubElement(custom_field, 'name')
                name.append(CDATA('HGI Req Verification Team'))
                value = ET.SubElement(custom_field, 'value')
                ver_team = item[REQ_VER_TEAM].split('\n')
                if len(ver_team) == 1:
                    ver_team = item[REQ_VER_TEAM].split(' ')
                if len(ver_team) == 1:
                    ver_team = item[REQ_VER_TEAM].split(',')
                if len(ver_team) == 1:
                    ver_team = item[REQ_VER_TEAM].split('|')
                if len(ver_team) == 1:
                    ver_team = item[REQ_VER_TEAM].split(';')
                ver_team = '|'.join(ver_team)
                value.append(CDATA(ver_team))

                if len(item) > REQ_COMMENT:
                    custom_field = ET.SubElement(custom_fields, 'custom_field')
                    name = ET.SubElement(custom_field, 'name')
                    name.append(CDATA('HGI Req Review Comments'))
                    value = ET.SubElement(custom_field, 'value')
                    value.append(CDATA(item[REQ_COMMENT]))
                    custom_field = ET.SubElement(custom_fields, 'custom_field')
                    name = ET.SubElement(custom_field, 'name')
                    name.append(CDATA('HGI Feature Phase'))
                    value = ET.SubElement(custom_field, 'value')
                    value.append(CDATA(item[REQ_PHASE]))


        if relation_list is not None:
            for relation_src in relation_list:
                for relation_dst in relation_src[1]:
                    relation = ET.SubElement(req_spec, 'relation')
                    source = ET.SubElement(relation, 'source')
                    source.text = relation_src[0]
                    destination = ET.SubElement(relation, 'destination')
                    destination.text = relation_dst
                    relation_type = ET.SubElement(relation, 'type')
                    relation_type.text = '1'

        rough_string = ET.tostring(tds, 'utf-8')
        #print rough_string
        reparsed = minidom.parseString(rough_string)
        f = open(filename, 'w')
        reparsed.writexml(f, newl='\n', encoding='utf-8')
        f.close()

        self.logger.info(self.log_prefix + \
                         "xml file %s was generated successfully." % \
                         (filename))
        return res

    def link_pfs2tds(self, tds_url, tc_url, pfs_url):
        tc_req_list = []
        req_tc_list = []
        res = None

        res = self.link_tc2tds(tds_url, tc_url, tc_req_list, req_tc_list)

        #pprint.pprint(tc_req_list)
        pfs_file = os.path.splitext(self.pfs_url)[0] + '.mm'
        tds_tc_file = self.tds_url.replace('.mm', '[TDS-TC].mm')
        res = self._build_fm_traceability(tds_tc_file, pfs_file, tc_req_list,
                                          self.tds_url.replace('.mm', '[TDS-TC-PFS].mm'))

        return res

    def link_tc2tds(self, tds_file, tc_file, tc_req_list=None, req_tc_list=None):
        if tc_req_list == None:
            tc_req_list = []
        if req_tc_list == None:
            req_tc_list = []

        tc_fm_file = tc_file.replace('.xml', '.mm')
        res = self._read_tc_from_xml(tc_file, tc_fm_file, tc_req_list)
        res = self._reverse_links(tc_req_list, req_tc_list)
        #pprint.pprint(req_tc_list)

        fm_tree = ET.parse(tds_file)
        fm_root = fm_tree.getroot()
        self._remove_node_prefix(fm_root)
        self._add_node_prefix(fm_root, '0')
        self._remove_link_node(fm_root)
        fm_tree.write(tds_file)

        res = self._build_fm_traceability(tds_file, tc_fm_file, req_tc_list, tds_file.replace('.mm', '[TDS-TC].mm'))

        return res


    def _link_tc_node(self, tc_id, tc_title, link_id, node):
        for child in node.iter('node'):
            # Make sure this is not the test case or requirement link node since only they are nodes with links
            if not child.attrib.has_key('LINK'):
                # If the TDS prefix matches with the link ID, then add this link to its sub-node
                if child.attrib['TEXT'].split(' ')[0] == link_id.split('_')[-1]:
                    link_text = 'HDVB-' + tc_id + ':' + tc_title
                    link_url = 'http://testlink.ea.mot.com/linkto.php?tprojectPrefix=HDVB&item=testcase&id=HDVB-' + tc_id
                    ET.SubElement(child, 'node', {'COLOR': '#990000', 'LINK': link_url, 'TEXT': link_text})
                    self.logger.info(self.log_prefix + \
                                     "Adding linkage sub-node (%s) to node (%s)" % \
                                     (link_text, child.attrib['TEXT']))
                    return 0

        self.logger.error(self.log_prefix + \
                          "Cannot find %s" % \
                          (tc_id))

        return None

    def _read_tc_from_xml(self, xml_file, fm_file, tc_req_list):
        tc_tree = lxmlET.parse(xml_file)
        tc_root = tc_tree.getroot()

        # Build the FreeMind for test case
        title = os.path.splitext(os.path.split(xml_file)[-1])[0]
        res = self._gen_tc_freemind(xml_file, title, fm_file)
        # Construct the traceability list between Test cases and Requirements/Test Design Specification  
        self.logger.info(self.log_prefix + \
                         "Getting traceability information from file %s" % \
                         (xml_file))
        prefix_list = [self.pmr_prefix, self.tds_prefix]
        prefix_list.extend(self.pfs_prefix.split(
            ','))  #Could be multiple PFS prefix since some requirements will be reused between projects.
        for tc in tc_root.iter('testcase'):
            req_links = []
            tc_id = self.repo_prefix + '-' + tc.find('externalid').text
            for req in tc.iter('requirement'):
                doc_id = req.find('doc_id').text

                for prefix in prefix_list:
                    # Check if this is a valid requirement/TDS for this project
                    if len(doc_id.split(prefix)) == 2:
                        req_links.append(doc_id.split(prefix)[1])
                        break
            # Please note the tc_id here is with the project prefix, and the req_id is without requirement prefix
            tc_req_list.append([tc_id, req_links])

        return res

    def _gen_tc_freemind(self, tc_file, title, output_file):
        ''' req_list is a list like [GROUP_NAME, [ [REQ_ID, REQ_TITLE, REQ_DESC, REQ_VER_TEAM], ... ] ]
            REQ_ID and REQ_TITLE will be combined as the node text and REQ_DESC will be displayed as comments
        '''
        tc_tree = lxmlET.parse(tc_file)
        tc_root = tc_tree.getroot()

        freemind = ET.Element('map', {'version': '1.0.1'})

        ET.SubElement(freemind, 'attribute_registry', {'SHOW_ATTRIBUTES': 'hide'})
        root_node = ET.SubElement(freemind, 'node', {'BACKGROUND_COLOR': '#0000ff', 'COLOR': '#000000', 'TEXT': title})
        ET.SubElement(root_node, 'font', {'NAME': 'SansSerif', 'SIZE': '20'})
        ET.SubElement(root_node, 'hook', {'NAME': 'accessories/plugins/AutomaticLayout.properties'})

        self._add_tc_details(tc_root, root_node)
        ET.ElementTree(freemind).write(output_file)
        self.logger.info(self.log_prefix + \
                         "Successfully generate test case FreeMind file %s" % \
                         (output_file))
        return 0

    def _add_tc_details(self, tc_root, fm_root):
        for child in tc_root:
            if child.tag == 'testsuite':
                #add a node in Freemind and call again.
                testsuite_node = ET.SubElement(fm_root, 'node',
                                               {'COLOR': '#990000', 'FOLDED': "true", 'TEXT': child.attrib['name']})
                self._add_tc_details(child, testsuite_node)
            if child.tag == 'testcase':
                #add a node in Freemind
                valid_tc = False
                node_comment = ''
                node_text = child.attrib['name']
                expected_results = ''
                tc_id = ''
                regression_level = ''
                for item in child:
                    if item.tag == 'externalid':
                        tc_id = item.text
                        node_text = self.repo_prefix + '-' + tc_id + PREFIX_TITLE_SEP + node_text
                    if item.tag == 'summary':
                        node_comment = '<p>Summary:</p>' + str(item.text) + '<p></p>'
                    if item.tag == 'preconditions':
                        node_comment = node_comment + '<p>Preconditions:</p>' + str(item.text) + '<p></p>'
                    if item.tag == 'steps':
                        node_comment = node_comment + '<p>Steps:</p>'
                        expected_results = '<p>Expected results:</p>'
                        for step in item.iter():
                            if step.tag == 'step_number':
                                node_comment = node_comment + '<p>' + step.text + '.'
                                expected_results = expected_results + '<p>' + step.text + '.'
                            if step.tag == 'actions':
                                node_comment = node_comment + str(step.text).replace('<p>', '', 1)
                            if step.tag == 'expected_results':
                                expected_results = expected_results + str(step.text).replace('<p>', '', 1)
                    if item.tag == 'custom_fields':
                        for custom_field in item:
                            if list(custom_field)[0].text == 'HGI Regression Level':
                                regression_level = 6 - len(list(custom_field)[1].text.split('|'))
                                #Enable this once the test case is linked with requirements
                                #                    if item.tag == 'requirements':
                                #                        for requirement in item:
                                #                            doc_id = list(requirement)[1].text
                                #                            for prefix in [self.pfs_prefix, self.pmr_prefix, self.tds_prefix]:
                                #                                if len(doc_id.split(prefix)) == 2:
                                #                                    valid_tc = True
                                #                                    break
                                #                if not valid_tc:
                                #                    continue
                node_comment = node_comment + '<p></p>' + expected_results
                node_link = self.testlink_url + '/linkto.php?tprojectPrefix=' + self.repo_prefix + '&item=testcase&id=' + self.repo_prefix + '-' + tc_id
                tc_node = ET.SubElement(fm_root, 'node', {'COLOR': '#990000', 'LINK': node_link, 'TEXT': node_text})
                richcontent = ET.SubElement(tc_node, 'richcontent', {'TYPE': 'NOTE'})
                html = ET.SubElement(richcontent, 'html')
                ET.SubElement(richcontent, 'head')
                body = ET.SubElement(html, 'body')
                for section in node_comment.replace('</p>', '').split('<p>'):
                    comment = ET.SubElement(body, 'p')
                    comment.text = section

                ET.SubElement(tc_node, 'icon', {'BUILTIN': 'full-' + str(regression_level)})
        return 0

    def link_tds2tc(self, fm_file, tc_file):
        test_suite = []
        tc_tds_list = []
        tc_pfs_list = []

        #res = self._read_tc_from_xml(tc_file, test_suite, tc_tds_list, tc_pfs_list)
        #res = self._gen_tc_freemind(test_suite)

        # res = self._link_tds_tc(self.tds_url)
        # res = self._link_tds_pfs()
        # res = self._link_pfs_tc()

        self.fm_file = fm_file
        tds_title = os.path.split(os.path.splitext(self.fm_file)[0])[1]
        self.fm_tree = lxmlET.parse(fm_file)
        fm_root = self.fm_tree.getroot()

        #parser = lxmlET.XMLParser(False)
        self.tc_file = tc_file
        self.tc_tree = ET.parse(tc_file)
        tc_root = self.tc_tree.getroot()

        # Firstly put all test cases with requirements/TDS links into a list
        link_list = []
        self._get_link_node(fm_root, link_list)
        #pprint.pprint(link_list)

        #Secondly loop through all test cases and add the TDS linkage in
        for tc in tc_root.iter('testcase'):
            tc_name = tc.get('name')
            tc_id = tc.find('externalid').text
            for tds_link in link_list:
                if tc_id == tds_link[0].split('-')[-1]:
                    if 1:
                        tds_link_found = False
                        for req in tc.iter('requirement'):
                            if (req.find('req_spec_title').text == tds_title) and \
                                    (req.find('doc_id').text.split('_')[-1] == tds_link[3]):
                                tds_link_found = True
                                break
                        if not tds_link_found:
                            requirements = tc.find('requirements')
                            if requirements == None:
                                requirements = ET.SubElement(tc, 'requirements')
                            link_item = ET.SubElement(requirements, 'requirement')
                            req_spec_title = ET.SubElement(link_item, 'req_spec_title')
                            #req_spec_title.text = lxmlET.CDATA(tds_title)
                            req_spec_title.append(CDATA(tds_title))
                            doc_id = ET.SubElement(link_item, 'doc_id')
                            doc_id.append(CDATA(tds_link[2]))
                            #                            self.logger.info(self.log_prefix + \
                            #                                "Add TDS link (%s) in test case (%s:%s)" % \
                            #                                (tds_link[3], tc_id, tc_name)

                            #        filename = os.path.splitext(self.tc_file)[0] + "_New.xml"
                            #        rough_string = ET.tostring(tc_root, 'utf-8')
                            #        reparsed = minidom.parseString(rough_string)
                            #        f= open(filename, 'w')
                            #        reparsed.writexml(f, addindent='  ', newl='\n',encoding='utf-8')
                            #        f.close()

        self.tc_tree.write(os.path.splitext(self.tc_file)[0] + "_New.xml")
        return 0

    #    def create_test_plan(self, tp_url, based_tp_url, auto_sync, ver_team):
    #        ''' The inputs could be TDS aided test planning, Test Suites aided test planning or PFS aided test planning.
    #
    #        '''
    #        removed_tc_list = []
    #        tc_list = []
    #        new_tc_list = []
    #        fm_tree = ET.parse(tp_url)
    #        tp_root = fm_tree.getroot()
    #        based_tp_root = ET.parse(based_tp_url).getroot()
    #        #Firstly we need to go through the based test plan to see if there any node or test cases are removed from the based test plan.
    #        res = self._find_removed_tc(based_tp_root, tp_root, removed_tc_list)
    #        pprint.pprint(removed_tc_list)
    #        #Secondly we need to go through the new test plan and remove all the test cases based on information above, regression levels and verification teams.
    #        res = self._update_tp(tp_root, ver_team, removed_tc_list)
    #        #Remove the nodes without any test case
    #        res = self._remove_node_wo_tc(tp_root)
    #        #Get test case list
    #        res = self._get_fm_tc_list(tp_root, tc_list)
    #        #Remove duplicate test cases in list
    #        res = self._remove_duplicate(tc_list, new_tc_list)
    #        pprint.pprint(new_tc_list)
    #        #Generate the test plan for importing to TestLink
    #        # platform, name, id, version, order
    #
    #        #Create the test plan
    #        if auto_sync == '1':
    #            tp_name = os.path.split(os.path.splitext(tp_url)[0])[-1]
    #            res = self._create_test_plan_in_tl(tp_name, new_tc_list)
    #
    #        return res

    def create_test_plan(self, tp_url, auto_sync, ver_team):
        ''' The inputs could be TDS aided test planning, Test Suites aided test planning or PFS aided test planning.                        
        '''
        removed_tc_list = []
        kept_tc_list = []
        tc_list = []
        new_tc_list = []
        fm_tree = ET.parse(tp_url)
        tp_root = fm_tree.getroot()

        #Firstly we need to go through the test plan to see if there any test case is removed or there are any test cases need to be kept.
        res = self._find_removed_kept_tc(tp_root, removed_tc_list, kept_tc_list)
        self.logger.info(self.log_prefix + \
                         "Test cases marked with remove icon are (%s)." % \
                         (removed_tc_list))
        self.logger.info(self.log_prefix + \
                         "Test cases marked with must-keep icon are (%s)." % \
                         (kept_tc_list))
        #Secondly we need to get all test cases based on information above, regression levels and verification teams.
        res = self._get_tc_list(tp_root, removed_tc_list, kept_tc_list, tc_list, ver_team)
        res = self._remove_duplicate(tc_list, new_tc_list)
        self.logger.info(self.log_prefix + \
                         "Test cases planned in this test cycle are (%s)." % \
                         (new_tc_list))

        #Update Test Plan
        res = self._update_fm_tp(tp_root, new_tc_list)
        fm_tree.write(tp_url)
        self.logger.info(self.log_prefix + \
                         "The original test plan file (%s) is updated." % \
                         (tp_url))

        #Generate the test plan for importing to TestLink
        # platform, name, id, version, order

        #Create the test plan
        if auto_sync == '1':
            tp_name = os.path.split(os.path.splitext(tp_url)[0])[-1]
            res = self._create_test_plan_in_tl(tp_name, new_tc_list)

        return res

    def _create_test_plan_in_tl(self, tp_name, tc_list):
        ''' Establish a connection with TestLink and then create a new test plan.
            Get the latest version the assigned test cases and then add them into the test plan.
            It could be very slow depending on the link and xmlrpc.
        '''
        self.logger.info(self.log_prefix + \
                         "Test plan (%s) will be created and updated in TestLink. This is going to take a while. Please wait..." % \
                         (tp_name))
        self.tls = testlink.TestLinkHelper().connect(testlink.TestlinkAPIClient)
        prj = self.tls.getTestProjectByName(self.repo_name)
        prj_id = prj['id']
        tp = self.tls.createTestPlan(tp_name, self.repo_name)
        tp_id = tp[0]['id']
        #tp_id = self.tls.getTestPlanByName(self.repo_name, tp_name)[0]['id']
        for tc_id in tc_list:
            tc_version = self.tls.getTestCase(None, testcaseexternalid=tc_id)[0]['version']
            self.tls.addTestCaseToTestPlan(prj_id, tp_id, tc_id, int(tc_version))

        self.logger.info(self.log_prefix + \
                         "Test plan (%s) is created and updated successfully." % \
                         (tp_name))

    def link_tp2tds_tc(self, tds_url, tc_url, name_filter):
        tc_list = []
        res = self._get_test_plan_info(name_filter, tc_list)
        #pprint.pprint(tc_list) 
        # Link TDS_TC file with Test Plan and Execution status
        #res = self.link_tc2tds(self.tds_url, self.tc_url)
        res = self._link_tp2fm(tds_url.replace('.mm', '[TDS-TC].mm'), tc_list)

    def _link_tp2fm(self, fm_file, tc_list):
        tp_list = []
        fm_tree = ET.parse(fm_file)
        root_node = fm_tree.getroot()
        for child in root_node.iter('node'):
            node_text = child.attrib['TEXT'].strip()
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            # If this is the node for a test case            
            if (tc_id.count(self.repo_prefix) == 1):
                tp_list = []
                for tc in tc_list:
                    if tc[0] == tc_id:
                        tp_list = tc[1]
                        break
                #print tp_list
                for tp in tp_list:
                    tp_name = tp[0]
                    tp_sts = tp[1]
                    tp_node = ET.SubElement(child, 'node', {'TEXT': tp_name})
                    if tp_sts == 'p':
                        ET.SubElement(tp_node, 'icon', {'BUILTIN': 'go'})
                    if tp_sts == 'f':
                        ET.SubElement(tp_node, 'icon', {'BUILTIN': 'stop'})
                    if tp_sts == 'b':
                        ET.SubElement(tp_node, 'icon', {'BUILTIN': 'prepare'})
                    if tp_sts == 'n':
                        ET.SubElement(tp_node, 'icon', {'BUILTIN': 'help'})
        fm_tree.write(fm_file.replace('.mm', '-TP.mm'))
        self.logger.info(self.log_prefix + \
                         "Successfully linked the test plan and execution results to file (%s)." % \
                         (fm_file.replace('.mm', '-TP.mm')))

    def _get_test_plan_info(self, name_filter, tc_list):
        self.logger.info(self.log_prefix + \
                         "Getting test plan and execution status from TestLink. This is going to take a while. Please wait...")
        self.tls = testlink.TestLinkHelper().connect(testlink.TestlinkAPIClient)
        prj = self.tls.getTestProjectByName(self.repo_name)
        prj_id = prj['id']
        tp_list = self.tls.getProjectTestPlans(prj_id)
        self.logger.info(self.log_prefix + \
                         "There are totally %d test plan for this project (%s)." % \
                         (len(tp_list), self.repo_name))
        for tp in tp_list:
            tp_name = tp['name']
            #TODO: Apply the name filter
            tp_id = tp['id']
            tc_dict = self.tls.getTestCasesForTestPlan(tp_id)
            for k in tc_dict.keys():
                tc = tc_dict[k][0]
                tc_id = tc['full_external_id']
                tc_sts = tc['exec_status']
                self._add_tc_history_list(tc_id, tc_sts, tp_name, tc_list)

        return 0

    def _add_tc_history_list(self, tc_id, tc_sts, tp_name, tc_list):
        for tc in tc_list:
            if tc_id == tc[0]:
                tc[1].append([tp_name, tc_sts])
                return True
        tc_list.append([tc_id, [[tp_name, tc_sts]]])
        return True

    def _remove_duplicate(self, old_list, new_list):
        for i in old_list:
            if not i in new_list:
                new_list.append(i)

    def _get_fm_tc_list(self, root_node, tc_list):
        for child in root_node.iter('node'):
            node_text = child.attrib['TEXT'].strip()
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            # If this is the node for a test case
            if tc_id.count(self.repo_prefix) == 1:
                tc_list.append(tc_id)

    def _update_fm_tp(self, root_node, tc_list):
        for child in root_node.findall('node'):
            for hook_node in child.findall('hook'):
                if hook_node.attrib['NAME'].strip() == 'accessories/plugins/AutomaticLayout.properties':
                    child.remove(hook_node)

            node_text = child.attrib['TEXT'].strip()
            #print node_text
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            # If this is the node for a test case
            if tc_id.count(self.repo_prefix) == 1:
                if tc_id in tc_list:
                    child.attrib['COLOR'] = '#000000'
                else:
                    child.attrib['COLOR'] = '#cccccc'

            if not self._has_valid_tc_node(child, tc_list):
                child.attrib['COLOR'] = '#cccccc'
                child.attrib['FOLDED'] = 'true'
            else:
                child.attrib['COLOR'] = '#000000'
                child.attrib['FOLDED'] = 'false'
            self._update_fm_tp(child, tc_list)

        return 0

    def _has_valid_tc_node(self, root_node, tc_list):
        for child in root_node.iter('node'):
            node_text = child.attrib['TEXT'].strip()
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            # If this is the node for a test case
            if (tc_id.count(self.repo_prefix) == 1) and (tc_id in tc_list):
                return True
        return False

    def _remove_node_wo_tc(self, root_node):
        for child in root_node.findall('node'):
            if not self._has_tc_node(child):
                root_node.remove(child)
            else:
                self._remove_node_wo_tc(child)

        return 0

    def _has_tc_node(self, root_node):
        for child in root_node.iter('node'):
            node_text = child.attrib['TEXT'].strip()
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            # If this is the node for a test case
            if tc_id.count(self.repo_prefix) == 1:
                return True
        return False

    def _get_tc_list(self, root_node, exclude_tc_list, kept_tc_list, tc_list, ver_team, regression_level='5'):
        for child in root_node.findall('node'):
            node_text = child.attrib['TEXT'].strip()
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            node_reg_lvl = regression_level
            for node_icon in child.findall('icon'):
                if node_icon.attrib['BUILTIN'].strip().count('full-') == 1:
                    node_reg_lvl = node_icon.attrib['BUILTIN'].strip()[-1]
            # If this is the node for a test case
            if tc_id.count(self.repo_prefix) == 1:
                # TODO: If we want to implement verification team, we need add this information in this node  
                # Keep the node if regression level is matched and not in the exclude_tc_list, or it's in the must keep list kept_tc_list    
                if ((tc_id not in exclude_tc_list) and (int(node_reg_lvl) <= int(regression_level))) \
                        or (tc_id in kept_tc_list):
                    tc_list.append(tc_id)
                    #child.attrib['COLOR'] = '#000000'          
                    #else:
                    #child.attrib['COLOR'] = '#cccccc'
            else:
                self._get_tc_list(child, exclude_tc_list, kept_tc_list, tc_list, ver_team, node_reg_lvl)

        return 0

    def _update_tp(self, root_node, ver_team, exclude_tc_list, regression_level='5'):
        for child in root_node.findall('node'):
            node_text = child.attrib['TEXT'].strip()
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            node_reg_lvl = regression_level
            for reg_lvl_icon in child.findall('icon'):
                if reg_lvl_icon.attrib['BUILTIN'].strip().count('full-') == 1:
                    node_reg_lvl = reg_lvl_icon.attrib['BUILTIN'].strip()[-1]
            # If this is the node for a test case
            if tc_id.count(self.repo_prefix) == 1:
                # TODO: If we want to implement verification team, we need add this information in this node
                #tc_ver_team = node_text.split(PREFIX_TITLE_SEP)[1].split('|')
                if (tc_id in exclude_tc_list) or (int(node_reg_lvl) > int(regression_level)):
                    root_node.remove(child)
                    print node_text
            else:
                self._update_tp(child, ver_team, exclude_tc_list, node_reg_lvl)

        return 0

    def _find_removed_kept_tc(self, root_node, removed_tc_list, kept_tc_list):
        for child in root_node.findall('node'):
            node_text = child.attrib['TEXT'].strip()
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            # If this is the node for a test case
            if tc_id.count(self.repo_prefix) == 1:
                for icon_node in child.findall('icon'):
                    # TODO: How about multiple icons?
                    if icon_node.attrib['BUILTIN'].strip() == 'button_cancel':
                        removed_tc_list.append(tc_id)
                    if icon_node.attrib['BUILTIN'].strip() == 'button_ok':
                        kept_tc_list.append(tc_id)
            else:
                self._find_removed_kept_tc(child, removed_tc_list, kept_tc_list)

    def _find_removed_tc(self, root_node, tp_root, removed_tc_list):
        for child in root_node.findall('node'):
            node_text = child.attrib['TEXT'].strip()
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            # If this is the node for a test case
            if tc_id.count(self.repo_prefix) == 1:
                parent_text = root_node.attrib['TEXT'].strip()
                self.node_found = False
                self._find_combined_node(parent_text, node_text, tp_root)
                if not self.node_found:
                    removed_tc_list.append(tc_id)
            else:
                self._find_removed_tc(child, tp_root, removed_tc_list)

    def _find_combined_node(self, parent_text, tc_text, root_node):
        for child in root_node.findall('node'):
            node_text = child.attrib['TEXT'].strip()
            tc_id = node_text.split(PREFIX_TITLE_SEP)[0]
            # If this is the node for a test case
            if tc_id.count(self.repo_prefix) == 1:
                # Check the node text to see if it matches thus we know if this node exist in new test plan
                if (root_node.attrib['TEXT'].strip() == parent_text) and (node_text == tc_text):
                    # Need a global variable for this, haven't find a new way to replace this.
                    self.node_found = True
            else:
                self._find_combined_node(parent_text, tc_text, child)
        return 0

    def _get_link_node(self, node, link_list):
        for child in node.findall('node'):
            if child.attrib.has_key('LINK'):
                tds_id = node.attrib['TEXT'].split(' ')[0]
                tc_id = child.attrib['TEXT'].split(':')[0]
                tc_title = ''.join(child.attrib['TEXT'].split(':')[1:])
                link_list.append([tc_id, tc_title, tds_id])
            else:
                self._get_link_node(child, link_list)
        return 0

    def _remove_node_prefix(self, node):
        for child in node.iter('node'):
            # Make sure this is not the test case or requirement link node since only they are nodes with links
            if not child.attrib.has_key('LINK'):
                # If the node text is started with a number, then we consider it having added prefix
                if child.attrib['TEXT'][0].isdigit:
                    # Since Unicode may also be considered as numbers, we need to make sure this is unicode or prefix
                    if (child.attrib['TEXT'].split(PREFIX_TITLE_SEP)[0] <> child.attrib['TEXT']):
                        self.logger.debug(self.log_prefix + \
                                          "Prefix of node (%s) has been removed" % \
                                          (child.attrib['TEXT']))
                        child.attrib['TEXT'] = ''.join(child.attrib['TEXT'].split(PREFIX_TITLE_SEP)[1:])

        return 0

    def _remove_link_node(self, node):
        '''The key here is to use findall method since it will create a new children list'''
        for child in node.findall('node'):
            if child.attrib.has_key('LINK'):
                self.logger.debug(self.log_prefix + \
                                  "Link node (%s) has been removed from parent node (%s)" % \
                                  (child.attrib['TEXT'], node.attrib['TEXT']))
                node.remove(child)

            else:
                self._remove_link_node(child)

    def _add_node_prefix(self, node, num):
        ''' Add the node prefix (something like 1.1.2.1) for the TDS document
        '''
        res = 0
        i = 0
        for child in node:
            if child.tag == 'node':
                i = i + 1
                prefix = str(num) + '.' + str(i)
                # Make sure this is not the test case or requirement link node since only they are nodes with links
                if not child.attrib.has_key('LINK'):
                    # If the node text is started with a number, then we consider it has already been added prefix
                    if child.attrib['TEXT'][0].isdigit:
                        # Since Unicode may also be considered as numbers, we need to make sure this is unicode or prefix
                        if (child.attrib['TEXT'].split(PREFIX_TITLE_SEP)[0] <> child.attrib['TEXT']):
                            if child.attrib['TEXT'].split(PREFIX_TITLE_SEP)[0] <> prefix[4:]:
                                self.logger.error(self.log_prefix + \
                                                  "The original prefix (%s) in node (%s) doesn't match with the actual prefix (%s). This could be caused by a wrong order/position of this node." % \
                                                  (
                                                      child.attrib['TEXT'].split(PREFIX_TITLE_SEP)[0],
                                                      child.attrib['TEXT'],
                                                      prefix[4:]))
                                return None
                                #child.attrib['TEXT'] = prefix[4:] + ' ' + ''.join(child.attrib['TEXT'].split(' ')[1:])                             
                        else:
                            child.attrib['TEXT'] = prefix[4:] + PREFIX_TITLE_SEP + child.attrib['TEXT']
                    else:
                        child.attrib['TEXT'] = prefix[4:] + PREFIX_TITLE_SEP + child.attrib['TEXT']
                res = self._add_node_prefix(child, prefix)

        return res

    def extract_requirements(self, excel_file, template):
        pmr_list = []
        pfs_list = []
        pfs_pmr_list = []
        pmr_pfs_list = []
        prefixed_pmr_pfs_list = []

        if not os.path.exists(excel_file):
            self.logger.error(self.log_prefix + \
                              "Cannot find the specified file (%s). Action aborted." % \
                              (excel_file))
            return None
        if template == 'HGI':
            res = self._read_req_from_xls_hgi(excel_file, pmr_list, pfs_list, pfs_pmr_list)
        elif template == 'KreaTV':
            res = self._read_req_from_xls_kreatv(excel_file, pmr_list, pfs_list, pfs_pmr_list)

        res = self._reverse_links(pfs_pmr_list, pmr_pfs_list)
        res = self._add_req_prefix(pmr_pfs_list, prefixed_pmr_pfs_list)

        # Get the filename without extension.
        title = os.path.splitext(os.path.split(self.pmr_url)[-1])[0]
        res = self._gen_req_xml(pmr_list, title, self.pmr_url, self.pmr_prefix)
        title = os.path.splitext(os.path.split(self.pfs_url)[-1])[0]
        res = self._gen_req_xml(pfs_list, title, self.pfs_url, self.pfs_prefix, prefixed_pmr_pfs_list)

        title = os.path.splitext(os.path.split(self.pmr_url)[-1])[0]
        res = self._gen_req_freemind(pmr_list, title, self.pmr_url.replace('.xml', '.mm'), self.pmr_prefix)
        title = os.path.splitext(os.path.split(self.pfs_url)[-1])[0]
        res = self._gen_req_freemind(pfs_list, title, self.pfs_url.replace('.xml', '.mm'), self.pfs_prefix)
        res = self._build_fm_traceability(self.pfs_url.replace('.xml', '.mm'), self.pmr_url.replace('.xml', '.mm'),
                                          pfs_pmr_list, self.pfs_url.replace('.xml', '[PFS-PMR].mm'))
        res = self._build_fm_traceability(self.pmr_url.replace('.xml', '.mm'), self.pfs_url.replace('.xml', '.mm'),
                                          pmr_pfs_list, self.pmr_url.replace('.xml', '[PMR-PFS].mm'))
        return res

    def _add_req_prefix(self, pmr_pfs_list, prefixed_pmr_pfs_list):
        for i, pmr_item in enumerate(pmr_pfs_list):
            prefixed_pmr_pfs_list.append([self.pmr_prefix + pmr_item[0], []])
            for pfs_item in pmr_item[1]:
                prefixed_pmr_pfs_list[i][1].append(self.pfs_prefix + pfs_item)
        return 0

    def _reverse_links(self, orig_list, reversed_list):
        ''' The original list is something like [PFS_ID, [PMR_ID1, PMRID2,...]].
            The reversed list is something like [PMR_ID, [PFS_ID1, PFS_ID2]]
        '''
        self.logger.debug(self.log_prefix + \
                          "Reversing the traceability links.")
        for orig_link in orig_list:
            src_id = orig_link[0]
            for link_id in orig_link[1]:
                if link_id == '':
                    continue
                reversed_link_exist = False
                i = 0
                for i, reversed_link in enumerate(reversed_list):
                    if reversed_link[0] == link_id:
                        reversed_link_exist = True
                        break
                if reversed_link_exist:
                    reversed_list[i][1].append(src_id)
                else:
                    reversed_list.append([link_id, [src_id]])

                    #pprint.pprint(pmr_pfs_list)
        return 0

    def _build_fm_traceability(self, dst_fm, src_fm, link_list, output_file):
        ''' This function is using to two FreeMind maps by using the traceability list in link_list[]
            link_list[] has the format of either [PFS_ID, [PMR_ID1, PMRID2,...]] or [PMR_ID, [PFS_ID1, PFS_ID2]] depends on 
            what's the destination FreeMind map.
        '''
        self.logger.info(self.log_prefix + \
                         "Building the FreeMind traceability file %s (Between %s and %s)." % \
                         (output_file, dst_fm, src_fm))
        dst_fm_tree = ET.parse(dst_fm)
        dst_fm_root = dst_fm_tree.getroot()
        src_fm_root = ET.parse(src_fm).getroot()
        new_added_nodes = []

        for dst_node in dst_fm_root.iter('node'):
            last_node = True
            for child in dst_node.findall('node'):
                self.logger.debug(self.log_prefix + \
                                  "This node (%s) has the child (%s) in file %s." % \
                                  (dst_node.attrib['TEXT'], child.attrib['TEXT'], dst_fm))
                last_node = False
                break
            self.logger.debug(self.log_prefix + \
                              "This node (%s) is the last node? %s." % \
                              (dst_node.attrib['TEXT'], last_node))
            if not last_node:
                continue
            # Please note the new added nodes will be looped through iter again so we need to ignore that by using new_added_nodes[]
            if dst_node.attrib['TEXT'] not in new_added_nodes:
                dst_id = dst_node.attrib['TEXT'].strip().split(PREFIX_TITLE_SEP)[0]
                traceability_links = []
                for traceability in link_list:
                    if dst_id == traceability[0]:
                        traceability_links = traceability[1]
                        break
                if (traceability_links == []) or (traceability_links == ['']):
                    # Highlight the node with traceability missing
                    self.logger.warning(self.log_prefix + \
                                        "Highlight the node (%s) with missing traceability for file %s." % \
                                        (dst_node.attrib['TEXT'].strip(), output_file))
                    dst_node.set('BACKGROUND_COLOR', '#ff0000')
                for link_id in traceability_links:
                    if link_id == '':
                        continue
                    link_found = False
                    for src_node in src_fm_root.iter('node'):
                        if (src_node.attrib['TEXT'].split(PREFIX_TITLE_SEP)[0] == link_id):
                            link_found = True
                            dst_node.append(src_node)
                            new_added_nodes.append(src_node.attrib['TEXT'])
                            self.logger.debug(self.log_prefix + \
                                              "Add link %s to %s." % \
                                              (link_id, dst_id))
                            break
                    if not link_found:
                        self.logger.warning(self.log_prefix + \
                                            "Cannot find link %s for %s for file %s." % \
                                            (link_id, dst_id, output_file))
                        # Highlight the node with traceability missing
                        self.logger.warning(self.log_prefix + \
                                            "Highlight the node (%s) with missing traceability for file %s." % \
                                            (dst_node.attrib['TEXT'].strip(), output_file))
                        dst_node.set('BACKGROUND_COLOR', '#ff0000')

        dst_fm_tree.write(output_file)

        self.logger.info(self.log_prefix + \
                         "Successfully built the FreeMind traceability file %s (Between %s and %s)." % \
                         (output_file, dst_fm, src_fm))

        return 0

    def _link_pfs_pmr(self, dst_fm, src_fm, link_list, output_file):
        ''' This function is using to link PMR FreeMind map and PFS FreeMind map by using the traceability list in link_list[]
            link_list[] has the format of either [PFS_ID, [PMR_ID1, PMRID2,...]] or [PMR_ID, [PFS_ID1, PFS_ID2]] depends on 
            what's the destination FreeMind map.
        '''
        dst_fm_tree = ET.parse(dst_fm)
        dst_fm_root = dst_fm_tree.getroot()
        src_fm_root = ET.parse(src_fm).getroot()
        new_added_nodes = []

        for dst_node in dst_fm_root.iter('node'):
            # Please note the new added nodes will be looped through iter again so we need to ignore that by using new_added_nodes[]
            if dst_node.attrib.has_key('LINK') and (dst_node.attrib['TEXT'] not in new_added_nodes):
                req_id = dst_node.attrib['TEXT'].split(PREFIX_TITLE_SEP)[0]
                req_links = []
                for req_trace in link_list:
                    if req_id == req_trace[0]:
                        req_links = req_trace[1]
                        break
                if (req_links == []) or (req_links == ['']):
                    # Highlight the node with traceability missing
                    self.logger.warning(self.log_prefix + \
                                        "Cannot find the requirement links for %s." % \
                                        (req_id))
                    dst_node.set('BACKGROUND_COLOR', '#ff0000')
                for req_link_id in req_links:
                    if req_link_id == '':
                        continue
                    link_found = False
                    for src_node in src_fm_root.iter('node'):
                        if src_node.attrib.has_key('LINK') and (
                                    src_node.attrib['TEXT'].split(PREFIX_TITLE_SEP)[0] == req_link_id):
                            link_found = True
                            dst_node.append(src_node)
                            new_added_nodes.append(src_node.attrib['TEXT'])
                            self.logger.info(self.log_prefix + \
                                             "Add requirement link %s to %s." % \
                                             (req_link_id, req_id))
                            break
                    if not link_found:
                        self.logger.error(self.log_prefix + \
                                          "Cannot find requirement link %s for %s." % \
                                          (req_link_id, req_id))

        dst_fm_tree.write(output_file)

        return 0

    def _gen_req_freemind(self, req_list, title, output_file, prefix):
        ''' req_list is a list like [GROUP_NAME, [ [REQ_ID, REQ_TITLE, REQ_DESC, REQ_VER_TEAM], ... ] ]
            REQ_ID and REQ_TITLE will be combined as the node text and REQ_DESC will be displayed as comments
        '''
        self.logger.info(self.log_prefix + \
                         "Generating the FreeMind file %s (Document Title: %s. Document ID Prefix: %s)." % \
                         (output_file, title, prefix))
        freemind = ET.Element('map', {'version': '1.0.1'})

        ET.SubElement(freemind, 'attribute_registry', {'SHOW_ATTRIBUTES': 'hide'})
        root_node = ET.SubElement(freemind, 'node', {'BACKGROUND_COLOR': '#0000ff', 'COLOR': '#000000', 'TEXT': title})
        ET.SubElement(root_node, 'font', {'NAME': 'SansSerif', 'SIZE': '20'})
        ET.SubElement(root_node, 'hook', {'NAME': 'accessories/plugins/AutomaticLayout.properties'})

        req_count = 0
        for group in req_list:
            group_node = ET.SubElement(root_node, 'node', {'COLOR': '#990000', 'FOLDED': "true", 'TEXT': group[0]})
            i = 0
            for i, req_item in enumerate(group[1]):
                node_text = req_item[REQ_ID] + PREFIX_TITLE_SEP + req_item[REQ_TITLE]
                node_comment = req_item[REQ_DESC]
                node_link = self.testlink_url + '/linkto.php?tprojectPrefix=' + self.repo_prefix + '&item=req&id=' + prefix + \
                            req_item[REQ_ID]
                req_node = ET.SubElement(group_node, 'node', {'COLOR': '#990000', 'LINK': node_link, 'TEXT': node_text})
                richcontent = ET.SubElement(req_node, 'richcontent', {'TYPE': 'NOTE'})
                html = ET.SubElement(richcontent, 'html')
                ET.SubElement(richcontent, 'head')
                body = ET.SubElement(html, 'body')
                comment = ET.SubElement(body, 'p')
                comment.text = node_comment
            i = i + 1
            req_count = req_count + i
            group_node.attrib['TEXT'] = group_node.attrib['TEXT'] + '[' + str(i) + ']'

        root_node.attrib['TEXT'] = root_node.attrib['TEXT'] + '[' + str(req_count) + ']'
        ET.ElementTree(freemind).write(output_file)
        self.logger.info(self.log_prefix + \
                         "Successfully generated the FreeMind file %s (Document Title: %s. Document ID Prefix: %s)." % \
                         (output_file, title, prefix))
        return 0

    def _read_req_from_xls_hgi(self, file_name, pmr_list, pfs_list, trace_list):
        """ This function will read a Excel and extract PMR, PFS and traceability out of it.
        """
        self.logger.info(self.log_prefix + \
                         "Reading requirements from file (%s). This is going to take a while. Please wait..." % \
                         file_name)

        if os.path.splitext(file_name)[-1] != '.xls':
            self.logger.error(self.log_prefix + \
                              "I am sorry that I can not parse this file. Please convert it to a xls file.")
            exit(-1)
        src_wb = open_workbook(file_name, on_demand=True, formatting_info=True)

        pfs_phase_col = -1
        pfs_ft_col = -1
        col_defined = False
        pmr_pfs_trace_list = []
        pmr_index_list = []
        pfs_index_list = []
        for s in src_wb.sheets():
            src_sheet = src_wb.sheet_by_name(s.name)
            if s.name.find('Specification') != -1:
                pmr_grp_id = 0
                pfs_grp_id = 0

                for i, cell in enumerate(src_sheet.col(0)):
                    if not col_defined:
                        for j in range(0, src_sheet.ncols):
                            if src_sheet.cell_value(i, j).strip() == 'PMR Index':
                                pmr_index_col = j
                            if src_sheet.cell_value(i, j).strip() == 'PMR Description':
                                pmr_desc_col = j
                            if src_sheet.cell_value(i, j).strip() == 'Index':
                                pfs_index_col = j
                            if src_sheet.cell_value(i, j).strip() == 'Category':
                                pfs_cat_col = j
                            if src_sheet.cell_value(i, j).strip() == 'Phase':
                                pfs_phase_col = j
                            if src_sheet.cell_value(i, j).strip() == 'Description':
                                pfs_desc_col = j
                            if src_sheet.cell_value(i, j).strip() == 'DEV':
                                pfs_dev_col = j
                            if src_sheet.cell_value(i, j).strip() == 'DVT':
                                pfs_dvt_col = j
                            if src_sheet.cell_value(i, j).strip() == 'SI&T':
                                pfs_sit_col = j
                                col_defined = True
                            if src_sheet.cell_value(i, j).strip() == 'FT':
                                pfs_ft_col = j
                            if src_sheet.cell_value(i, j).strip().lower().endswith('comments'):
                                pmr_cmt_col = j
                    else:
                        pmr_index = src_sheet.cell_value(i, pmr_index_col).strip()
                        pmr_desc = src_sheet.cell_value(i, pmr_desc_col).strip()
                        pmr_ver_team = 'ATP'
                        if pmr_index != '' and pmr_desc == '':
                            # This is a PMR category
                            pmr_grp_desc = src_sheet.cell_value(i, pmr_index_col).strip()
                            pmr_list.append([pmr_grp_desc, []])
                            pmr_grp_id = len(pmr_list) - 1
                        if len(pmr_list) == 0:
                            pmr_list.append(['Default Category', []])

                        pfs_index = src_sheet.cell_value(i, pfs_index_col).strip()
                        pfs_desc = src_sheet.cell_value(i, pfs_desc_col).strip()
                        pfs_cat = src_sheet.cell_value(i, pfs_cat_col).strip()
                        pfs_dev = src_sheet.cell_value(i, pfs_dev_col).strip()
                        pfs_dvt = src_sheet.cell_value(i, pfs_dvt_col).strip()
                        pfs_sit = src_sheet.cell_value(i, pfs_sit_col).strip()
                        pmr_cmt = src_sheet.cell_value(i, pmr_cmt_col).strip()
                        if pmr_cmt != '':
                            pmr_cmt = 'SE Comments:' + pmr_cmt

                        pfs_ft = ''
                        if pfs_ft_col != -1:
                            # This is an optional column
                            pfs_ft = src_sheet.cell_value(i, pfs_ft_col).strip()
                        pfs_phase = ''
                        if pfs_phase_col != -1:
                            # This is an optional column
                            pfs_phase = str(src_sheet.cell_value(i, pfs_phase_col)).strip()
                            if not pfs_phase.upper().startswith('P'):
                                pfs_phase = 'P' + pfs_phase
                                pfs_phase = pfs_phase[:2]

                        for merged_range in src_sheet.merged_cells:
                            rlo, rhi, clo, chi = merged_range
                            if (i >= rlo) and (i < rhi) and (pfs_index_col >= clo) and (pfs_index_col < chi):
                                pfs_index = src_sheet.cell_value(rlo, pfs_index_col).strip()
                            if (i >= rlo) and (i < rhi) and (pfs_desc_col >= clo) and (pfs_desc_col < chi):
                                pfs_desc = src_sheet.cell_value(rlo, pfs_desc_col).strip()
                            if (i >= rlo) and (i < rhi) and (pmr_index_col >= clo) and (pmr_index_col < chi):
                                pmr_index = src_sheet.cell_value(rlo, pmr_index_col).strip()
                            if (i >= rlo) and (i < rhi) and (pmr_desc_col >= clo) and (pmr_desc_col < chi):
                                pmr_desc = src_sheet.cell_value(rlo, pmr_desc_col).strip()
                            if (i >= rlo) and (i < rhi) and (pfs_cat_col >= clo) and (pfs_cat_col < chi):
                                pfs_cat = src_sheet.cell_value(rlo, pfs_cat_col).strip()

                        if pmr_index == 'PMR Index':
                            continue

                        if pfs_cat != '':
                            pfs_cat_exist = False
                            for item_index, pfs_item in enumerate(pfs_list):
                                if pfs_cat == pfs_item[0]:
                                    # This is an existing PFS category
                                    pfs_grp_id = item_index
                                    pfs_cat_exist = True
                                    break
                            if not pfs_cat_exist:
                                # This is a new PFS category
                                pfs_list.append([pfs_cat, []])
                                pfs_grp_id = len(pfs_list) - 1
                        # else:
                        #     if len(pfs_list) != 0:
                        #         pfs_grp_id = len(pfs_list) - 1
                        #         pfs_cat = pfs_list[pfs_grp_id][0]

                        pfs_ver_team = ''
                        if pfs_dev.upper() == 'Y':
                            pfs_ver_team = '|DEV'
                        if pfs_dvt.upper() == 'Y':
                            pfs_ver_team += '|DVT'
                        if pfs_sit.upper() == 'Y':
                            pfs_ver_team += '|SIT'
                        if pfs_ft.upper() == 'Y':
                            pfs_ver_team += '|FT'
                        pfs_ver_team = '|'.join(pfs_ver_team.split('|')[1:])

                        if pmr_index != '' and pmr_desc != '' and pfs_index != '':
                            # PFS item traced to PMR item
                            if pmr_index not in pmr_index_list:
                                pmr_list[pmr_grp_id][1].append(
                                    [pmr_index, pmr_index, pmr_desc, pmr_ver_team, pmr_cmt, ''])
                                pmr_index_list.append(pmr_index)
                            if pfs_index not in pfs_index_list:
                                pfs_list[pfs_grp_id][1].append(
                                    [pfs_index, pfs_cat, pfs_desc, pfs_ver_team, '', pfs_phase])
                                pfs_index_list.append(pfs_index)
                            self._add_traceability(pmr_pfs_trace_list, pmr_index, [pfs_index])
                        if pmr_index == '' and pmr_desc == '' and pfs_index != '' and pfs_desc != '':
                            # New PFS item traced to previous PMR item
                            pmr_index = pre_pmr_index
                            if pfs_index not in pfs_index_list:
                                pfs_list[pfs_grp_id][1].append(
                                    [pfs_index, pfs_cat, pfs_desc, pfs_ver_team, '', pfs_phase])
                                pfs_index_list.append(pfs_index)
                            self._add_traceability(pmr_pfs_trace_list, pmr_index, [pfs_index])
                        if pmr_index == '' and pmr_desc == '' and pfs_index == '' and pfs_desc != '':
                            # Traceability only PFS item traced to previous PMR item
                            pmr_index = pre_pmr_index
                            self._add_traceability(pmr_pfs_trace_list, pmr_index, pfs_desc.split('\n'))
                        if pmr_index != '' and pmr_desc != '' and pfs_index == '' and pfs_desc != '':
                            # Existing PFS item traced to new PMR item
                            if pmr_index not in pmr_index_list:
                                pmr_list[pmr_grp_id][1].append(
                                    [pmr_index, pmr_index, pmr_desc, pmr_ver_team, pmr_cmt, ''])
                                pmr_index_list.append(pmr_index)
                            self._add_traceability(pmr_pfs_trace_list, pmr_index, pfs_desc.split('\n'))
                        if pmr_index != '' and pmr_desc != '' and pfs_index == '' and pfs_desc == '':
                            # New PMR item with no PFS item
                            # pfs_index = pre_pfs_index
                            if pmr_index not in pmr_index_list:
                                pmr_list[pmr_grp_id][1].append(
                                    [pmr_index, pmr_index, pmr_desc, pmr_ver_team, pmr_cmt, ''])
                                pmr_index_list.append(pmr_index)
                                # self._add_traceability(pmr_pfs_trace_list, pmr_index, pfs_index)

                        if pmr_index != '':
                            pre_pmr_index = pmr_index
                        if pfs_index != '':
                            pre_pfs_index = pfs_index

        self._reverse_links(pmr_pfs_trace_list, trace_list)
        #pprint.pprint(trace_list)
        self.logger.info(self.log_prefix + \
                         "Successfully extracted requirements from file (%s)." % \
                         (file_name))
        return 0

    def _add_traceability(self, trace_list, dst_index, src_index_list):
        """
        This function is used to generate a traceablity list like [PMR, [PFS1, PFS2, PFS3]]
        :param trace_list:
        :param dst_index:
        :param src_index_list:
        """
        for trace_item in trace_list:
            if dst_index == trace_item[0]:
                for new_src_index in src_index_list:
                    duplicated_src = False
                    for orig_src_index in trace_item[1]:
                        if new_src_index == orig_src_index:
                            duplicated_src = True
                            break
                    if not duplicated_src:
                        trace_item[1].append(new_src_index)
                return
        trace_list.append([dst_index, src_index_list])

    def _read_req_from_xls_kreatv(self, file_name, pmr_list, pfs_list, trace_list):
        ''' This function will read a Excel and extract PMR, PFS and traceability out of it.
        '''
        self.logger.info(self.log_prefix + \
                         "Reading requirements from file (%s). This is going to take a while. Please wait..." % \
                         (file_name))
        src_wb = open_workbook(file_name, on_demand=True)

        for s in src_wb.sheets():
            src_sheet = src_wb.sheet_by_name(s.name)
            if s.name == 'PMR':
                group_id = 0
                for i, cell in enumerate(src_sheet.col(0)):
                    req_id = cell.value.strip()
                    req_title = src_sheet.cell_value(i, 1).strip()
                    req_desc = src_sheet.cell_value(i, 2).strip()
                    ver_team = 'ATP'
                    if req_desc == '':
                        group_id = group_id + 1
                        pmr_list.append([req_title, []])
                    else:
                        pmr_list[group_id - 1][1].append([req_id, req_title, req_desc, ver_team])
                        #pprint.pprint(pmr_list)
            if s.name == 'Requirements':
                group_id = 0
                for i, cell in enumerate(src_sheet.col(0)):
                    if i > 0:
                        req_id = cell.value.strip()
                        req_title = src_sheet.cell_value(i, 1).strip()
                        ver_team = src_sheet.cell_value(i, 3).strip()
                        req_desc = src_sheet.cell_value(i, 4).strip()
                        if req_desc == '':
                            group_id = group_id + 1
                            pfs_list.append([req_id, []])
                        else:
                            pfs_list[group_id - 1][1].append([req_id, req_title, req_desc, ver_team])
                            #pprint.pprint(pfs_list)
            if s.name == 'PFS':
                group_id = 0
                for i, cell in enumerate(src_sheet.col(0)):
                    if i > 0:
                        req_id = cell.value.strip()
                        req_trace = src_sheet.cell_value(i, 2).strip().split('\n')
                        if len(req_trace) == 1:
                            req_trace = src_sheet.cell_value(i, 2).strip().split(' ')
                        if len(req_trace) == 1:
                            req_trace = src_sheet.cell_value(i, 2).strip().split(',')
                        if len(req_trace) == 1:
                            req_trace = src_sheet.cell_value(i, 2).strip().split(';')
                            #req_trace = '|'.join(req_trace)
                        if str(src_sheet.cell_value(i, 1)).strip() <> '':
                            trace_list.append([req_id, req_trace])
                            #pprint.pprint(trace_list)
        self.logger.info(self.log_prefix + \
                         "Successfully extracted requirements from file (%s)." % \
                         (file_name))
        return 0


def args_parser(arguments=None):
    parser = argparse.ArgumentParser(description= \
                                         'This application can be used to extract event test case, sub-procedure test cases and\
        multi-procedure test cases information from a Excel file, and generate test cases based \
        on these configurations. It can also be used to parse C header files to get the struct\
        template and message map.')
    parser.add_argument('--version', action='version', version='%(prog)s 0.1')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-ap', '--add_prefix', action='store_true',
                       help="Update the existing Freemind file and add the prefix to each node except for \
              nodes with links (for instance: test case node or requirement node). \
              The most common usage is FreeMind -ap -f FREEMIND_FILE.")
    group.add_argument('-rp', '--remove_prefix', action='store_true',
                       help="Update the existing Freemind file and remove the prefix to each node. \
              The most common usage is FreeMind -rp -f FREEMIND_FILE.")
    group.add_argument('-g', '--gen_tds', action='store_true',
                       help="Update the existing Freemind file and remove the prefix to each node. \
              The most common usage is FreeMind -rp -f FREEMIND_FILE.")
    group.add_argument('-l', '--link_tds', action='store_true',
                       help="Extract test case and TDS linkage information from xml file exported from TestLink.\
                and update the FreeMind file with test cases links.\
                The most common usage is FreeMind -l -f FREEMIND_FILE -xml XML_FILE.")

    parser.add_argument('-s', '--src_file',
                        help="Specify the FreeMind file which contains various nodes of test design specification.")

    parser.add_argument('-d', '--dst_file',
                        help="Specify the xml file exported from TestLink with test case and TDS linkage information.")

    if arguments == None:
        args = parser.parse_args()
    else:
        args = parser.parse_args(arguments)

    return args


def start_main():
    reload(sys)
    sys.setdefaultencoding('utf-8')
    logging.config.fileConfig(PKG_PATH + 'logging.conf')
    logger = logging.getLogger(__name__)
    cfg_file = './config.xml'
    if os.path.exists(cfg_file):
        FreeMind(logger, cfg_file)
        sys.exit()

    freemind = FreeMind(logger)
    args = args_parser()
    if (args.add_prefix and args.src_file != None):
        freemind.add_prefix(args.src_file)
        sys.exit()
    if (args.remove_prefix and args.src_file != None):
        freemind.remove_prefix(args.src_file)
        sys.exit()
    if (args.gen_tds and args.src_file != None):
        freemind.gen_tds(args.src_file)
        sys.exit()
    if (args.link_tds and args.src_file != None and args.dst_file != None):
        if os.path.splitext(args.src_file)[-1] == 'mm' and os.path.splitext(args.dst_file)[-1] == 'xml':
            freemind.link_tds2tc(args.src_file, args.dst_file)
        if os.path.splitext(args.src_file)[-1] == 'xml' and os.path.splitext(args.dst_file)[-1] == 'mm':
            freemind.link_tc2tds(args.dst_file, args.src_file)
        sys.exit()


if __name__ == '__main__':
    start_main()