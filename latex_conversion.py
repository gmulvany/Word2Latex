#BY GERARD MUULVANY
import os
import tempfile
from tkinter.font import BOLD
import pypandoc 
import re
from docx import Document
from lxml import etree
from xml.etree.ElementTree import QName


####################### GLOBAL PROPERTIES ############################
current_directory = os.getcwd()
                                                                         
doc_path = current_directory

# ANSI escape codes for console red text
bold_red = "\033[1;31m"
green = "\033[32m"
yellow = "\033[93m"
italic = "\033[3m"
reset = "\033[0m"

#Pandoc arguments
extra_args = [
               '--filter',
               'pandoc-fignos',
               '--citeproc',
               '--csl',
               'custom-key.csl',
               '--biblatex',
               '--bibliography',
               'Bibliography.bib',
               '--wrap',
               'none',
               '--extract-media',
               './','--verbose']
to_format = 'tex'
input_format = 'docx+citations'

acronyms =  {
                'Technosozia'                   : '\\myac{Technosozia}',
                'technosozia'                   : '\\myac{Technosozia}',

                'Technostartistry'              : '\\myac{Technostartistry}',
                'technostartistry'              : '\\myac{Technostartistry}',

                ' AI '                          : ' \\myac{AI} ',
                ' AI.'                          : ' \\myac{AI} ',
                ' AI,'                          : ' \\myac{AI} ',
                'Artificial Intelligences (AI)' : ' \\myacp{AI} ',
                'Artificial intelligences (AI)' : ' \\myacp{AI} ',
                'artificial intelligences (AI)' : ' \\myacp{AI} ',
                'Artificial Intelligences'      : ' \\myacp{AI} ',
                'Artificial intelligences'      : ' \\myacp{AI} ',
                'artificial intelligences'      : ' \\myacp{AI} ',
                'Artificial Intelligence (AI)'  : ' \\myac{AI} ',
                'Artificial intelligence (AI)'  : ' \\myac{AI} ',
                'artificial intelligence (AI)'  : ' \\myac{AI} ',
                'Artificial Intelligence'       : ' \\myac{AI} ',
                'Artificial intelligence'       : ' \\myac{AI} ',
                'artificial intelligence'       : ' \\myac{AI} ',


                ' AR '                          : ' \\myac{AR} ',
                ' AR.'                          : ' \\myac{AR} ',
                ' AR,'                          : ' \\myac{AR} ',
                'Augmented Reality (AR)'        : ' \\myac{AR} ',
                'Augmented reality (AR)'        : ' \\myac{AR} ',
                'augmented reality (AR)'        : ' \\myac{AR} ',
                'Augmented Reality'             : ' \\myac{AR} ',
                'augmented reality'             : ' \\myac{AR} ',
                'Augmented reality'             : ' \\myac{AR} ',

                ' AV '                          : ' \\myac{AV} ',
                ' AV.'                          : ' \\myac{AV} ',
                ' AV,'                          : ' \\myac{AV} ',
                'Augmented Virtuality (AV)'     : ' \\myac{AV} ',
                'Augmented virtuality (AV)'     : ' \\myac{AV} ',
                'augmented virtuality (AV)'     : ' \\myac{AV} ',
                'Augmented Virtuality'          : ' \\myac{AV} ',
                'augmented virtuality'          : ' \\myac{AV} ',
                'Augmented virtuality'          : ' \\myac{AV} ',    

                ' BTN '                         : ' \\myac{BTN} ',
                ' BTN.'                         : ' \\myac{BTN} ',
                ' BTN,'                         : ' \\myac{BTN} ',
                '"Because the Night"'           : ' \\myacl{BTN} ',
                '"Because the Night"'           : ' \\myacl{BTN} ',
                '``Because the Night"'          : ' \\myacl{BTN} ',
                '\'Because the Night\''         : ' \\myacl{BTN} ',
                '"Because The Night"'           : ' \\myacl{BTN} ',
                '"Because The Night"'           : ' \\myacl{BTN} ',
                '``Because The Night"'          : ' \\myacl{BTN} ',
                '\'Because The Night\''         : ' \\myacl{BTN} ',
                'Because The Night'           : ' \\myacl{BTN} ',
                'Because the Night'           : ' \\myacl{BTN} ',

                ' CAD '                         : ' \\myac{CAD} ',
                ' CAD.'                         : ' \\myac{CAD} ',
                ' CAD,'                         : ' \\myac{CAD} ',
                'Computer-Aided Design (CAD)'   : ' \\myac{CAD} ',
                'Computer Aided Design'         : ' \\myac{CAD} ',
                'Computer-Aided Design'         : ' \\myac{CAD} ',
                'Computer-aided design'         : ' \\myac{CAD} ',
                'computer-aided design'         : ' \\myac{CAD} ',
                
                ' CCC '                             : ' \\myac{CCC} ',
                ' CCC.'                             : ' \\myac{CCC} ',
                ' CCC,'                             : ' \\myac{CCC} ',
                'Content Community Complex (CCC)'   : ' \\myac{CCC} ',
                'Content community complex (CCC)'   : ' \\myac{CCC} ',
                'content community complex (CCC)'   : ' \\myac{CCC} ',
                'Content Community Complex'         : ' \\myac{CCC} ',
                'Content community complex'         : ' \\myac{CCC} ',
                'content community complex'         : ' \\myac{CCC} ',
                'Content-Community Complex (CCC)'   : ' \\myac{CCC} ',
                'Content-community complex (CCC)'   : ' \\myac{CCC} ',
                'content-community complex (CCC)'   : ' \\myac{CCC} ',
                'Content-Community Complex'         : ' \\myac{CCC} ',
                'Content-community complex'         : ' \\myac{CCC} ',
                'content-community complex'         : ' \\myac{CCC} ',

                ' CV '                          : ' \\myac{CV} ',
                ' CV.'                          : ' \\myac{CV} ',
                ' CV,'                          : ' \\myac{CV} ',
                'Computer Vision (CV)'     : ' \\myac{CV} ',
                'Computer vision (CV)'     : ' \\myac{CV} ',
                'computer vision (CV)'     : ' \\myac{CV} ',
                'Computer Vision'          : ' \\myac{CV} ',
                'Computer vision'          : ' \\myac{CV} ',
                'computer vision'          : ' \\myac{CV} ',
                
                ' DSR '                         : ' \\myac{DSR} ',
                ' DSR.'                         : ' \\myac{DSR} ',
                ' DSR,'                         : ' \\myac{DSR} ',
                'Double System Recording (DSR)' : ' \\myac{DSR} ',
                'Double System Recording'       : ' \\myac{DSR} ',
                'Double system recording'       : ' \\myac{DSR} ',
                'double system recording'       : ' \\myac{DSR} ',
                                                   
                ' DTNC '                                : ' \\myac{DTNC} ',
                ' DTNC.'                                : ' \\myac{DTNC} ',
                ' DTNC,'                                : ' \\myac{DTNC} ',
                'Digital Twins Native Continuum (DTNC)' : ' \\myac{DTNC} ',
                'Digital Twins Native Continuum'        : ' \\myac{DTNC} ',
                'Digital twins native continuum'        : ' \\myac{DTNC} ',
                'digital twins native continuum'        : ' \\myac{DTNC} ',
                
                ' EMS '                             : ' \\myac{EMS} ',
                ' EMS.'                             : ' \\myac{EMS} ',
                ' EMS,'                             : ' \\myac{EMS} ',
                'Electro-Muscle Stimulation (EMS)'        : ' \\myAc{EMS} ',
                'Electro-muscle stimulation (EMS)'        : ' \\myAc{EMS} ',
                'electro-muscle stimulation (EMS)'        : ' \\myac{EMS} ',
                'electro-muscle stimulation'              : ' \\myac{EMS} ',
                'electro-muscle stimulation'              : ' \\myAc{EMS} ',
                'electro-muscle stimulation'              : ' \\myac{EMS} ',
                'electro-muscle stimulation (EMS)'        : ' \\myAc{EMS} ',
                'electro-muscle stimulation (EMS)'        : ' \\myAc{EMS} ',
                'electro-muscle stimulation (EMS)'        : ' \\myac{EMS} ',
                'electro-muscle stimulation'              : ' \\myac{EMS} ',
                'electro-muscle stimulation'              : ' \\myAc{EMS} ',
                'electro-muscle stimulation'              : ' \\myac{EMS} ',

                ' FK '                          : ' \\myac{FK} ',
                ' FK.'                          : ' \\myac{FK} ',
                ' FK,'                          : ' \\myac{FK} ',
                'Forward Kinematics (IK)'       : ' \\myac{FK} ',
                'Forward Kinematics'            : ' \\myac{FK} ',
                'Forward kinematics'            : ' \\myac{FK} ',
                'forward kinematics'            : ' \\myac{FK} ',

                ' HCI '                             : ' \\myac{HCI} ',
                ' HCI.'                             : ' \\myac{HCI} ',
                ' HCI,'                             : ' \\myac{HCI} ',
                'Human Computer Interaction (HCI)'  : ' \\myac{HCI} ',
                'Human Computer Interaction'        : ' \\myac{HCI} ',
                'Human computer interaction'        : ' \\myac{HCI} ',
                'human computer interaction'        : ' \\myac{HCI} ',


                ' HMD '                             : ' \\myac{HMD} ',
                ' HMD.'                             : ' \\myac{HMD} ',
                ' HMD,'                             : ' \\myac{HMD} ',                 
                ' HMDs '                             : ' \\myacp{HMD} ',
                ' HMDs.'                             : ' \\myacp{HMD} ',
                ' HMDs,'                             : ' \\myacp{HMD} ',
                'head mounted displays (HMDs)'        : ' \\myacp{HMD} ',
                'head-mounted displays'              : ' \\myacp{HMD} ',
                'head mounted displays'              : ' \\myacp{HMD} ',
                'head-Mounted displays (HMDs)'        : ' \\myacp{HMD} ',
                'Head Mounted Display (HMD)'        : ' \\myAc{HMD} ',
                'Head mounted display (HMD)'        : ' \\myAc{HMD} ',
                'head mounted display (HMD)'        : ' \\myac{HMD} ',
                'Head Mounted Display'              : ' \\myac{HMD} ',
                'Head mounted display'              : ' \\myAc{HMD} ',
                'head mounted display'              : ' \\myac{HMD} ',
                'Head-Mounted Display (HMD)'        : ' \\myAc{HMD} ',
                'Head-Mounted display (HMD)'        : ' \\myAc{HMD} ',
                'head-Mounted display (HMD)'        : ' \\myac{HMD} ',
                'Head-Mounted Display'              : ' \\myac{HMD} ',
                'Head-mounted display'              : ' \\myAc{HMD} ',
                'head-mounted display'              : ' \\myac{HMD} ',


                ' IoT '                             : ' \\myac{IoT} ',
                ' IoT.'                             : ' \\myac{IoT} ',
                ' IoT,'                             : ' \\myac{IoT} ',
                ' IOT '                             : ' \\myac{IoT} ',
                ' IOT.'                             : ' \\myac{IoT} ',
                ' IOT,'                             : ' \\myac{IoT} ',
                'Internet of Things (IoT)'  : ' \\myAc{IoT} ',
                'Internet Of Things (IoT)'  : ' \\myAc{IoT} ',
                'Internet of things (IoT)'  : ' \\myAc{IoT} ',
                'internet of things (IoT)'  : ' \\myac{IoT} ',
                'Internet of Things'        : ' \\myac{IoT} ',
                'Internet of things'        : ' \\myAc{IoT} ',
                'internet of things'        : ' \\myac{IoT} ',
                'Internet-of-Things (IoT)'  : ' \\myAc{IoT} ',
                'Internet-Of-Things (IoT)'  : ' \\myAc{IoT} ',
                'Internet-of-things (IoT)'  : ' \\myAc{IoT} ',
                'internet-of-things (IoT)'  : ' \\myac{IoT} ',
                'Internet-of-Things'        : ' \\myac{IoT} ',
                'Internet-of-things'        : ' \\myAc{IoT} ',
                'internet-of-things'        : ' \\myac{IoT} ',

                ' IK '                          : ' \\myac{IK} ',
                ' IK.'                          : ' \\myac{IK} ',
                ' IK,'                          : ' \\myac{IK} ',
                'Inverse Kinematics (IK)'       : ' \\myac{IK} ',
                'Inverse kinematics (IK)'       : ' \\myAc{IK} ',
                'inverse kinematics (IK)'       : ' \\myac{IK} ',
                'Inverse Kinematics'            : ' \\myac{IK} ',
                'Inverse kinematics'            : ' \\myAc{IK} ',
                'inverse kinematics'            : ' \\myac{IK} ',

                ' IR '                          : ' \\myac{IR} ',
                ' IR.'                          : ' \\myac{IR} ',
                ' IR,'                          : ' \\myac{IR} ',
                'Infra-Red (IR)'                : ' \\myac{IR} ',
                'Infra Red (IR)'                : ' \\myac{IR} ',
                'Infrared (IR)'                : ' \\myac{IR} ',
                'Infra-Red'                     : ' \\myac{IR} ',
                'Infra-red'                     : ' \\myAc{IR} ',
                'infra-red'                     : ' \\myac{IR} ',
                'Infra Red'                     : ' \\myac{IR} ',
                'Infra red'                     : ' \\myAc{IR} ',
                'infra red'                     : ' \\myac{IR} ',
                'Infrared'                     : ' \\myac{IR} ',
                'infrared'                     : ' \\myac{IR} ',

                ' MR '                          : ' \\myac{MR} ',
                ' MR.'                          : ' \\myac{MR} ',
                ' MR,'                          : ' \\myac{MR} ',
                'Mixed Reality (MR)'            : ' \\myac{MR} ',
                'Mixed reality (MR)'            : ' \\myAc{MR} ',
                'mixed reality (MR)'            : ' \\myac{MR} ',
                'Mixed Reality'                 : ' \\myac{MR} ',
                'Mixed reality'                 : ' \\myAc{MR} ',
                'mixed reality'                 : ' \\myac{MR} ',
                
                ' NLE '                         : ' \\myac{NLE} ',
                ' NLE.'                         : ' \\myac{NLE} ',
                ' NLE,'                         : ' \\myac{NLE} ',
                ' NLEs '                         : ' \\myacp{NLE} ',
                ' NLEs.'                         : ' \\myacp{NLE} ',
                ' NLEs,'                         : ' \\myacp{NLE} ',
                'Non-Linear Editor (NLE)'       : ' \\myac{NLE} ',
                'Non-linear editor (NLE)'       : ' \\myAc{NLE} ',
                'non-linear editor (NLE)'       : ' \\myac{NLE} ',
                'Non Linear Editor (NLE)'       : ' \\myac{NLE} ',
                'Non linear editor (NLE)'       : ' \\myac{NLE} ',
                'non linear editor (NLE)'       : ' \\myac{NLE} ',
                'Non-Linear Editors (NLE)'       : ' \\myAcp{NLE} ',
                'Non-linear editors (NLE)'       : ' \\myAcp{NLE} ',
                'non-linear editors (NLE)'       : ' \\myacp{NLE} ',
                'Non Linear Editors (NLE)'       : ' \\myacp{NLE} ',
                'Non linear editors (NLE)'       : ' \\myacp{NLE} ',
                'non linear editor (NLE)'       : ' \\myac{NLE} ',
                'Non-Linear Editor'             : ' \\myac{NLE} ',
                'Non-linear editor'             : ' \\myac{NLE} ',
                'non-linear editor'             : ' \\myac{NLE} ',
                'Non Linear Editor'             : ' \\myac{NLE} ',
                'Non linear editor'             : ' \\myAc{NLE} ',
                'non linear editor'             : ' \\myac{NLE} ',

                ' RE '                          : ' \\myac{RE} ',
                ' RE.'                          : ' \\myac{RE} ',
                ' RE,'                          : ' \\myac{RE} ',
                'Real Environments (RE)'         : ' \\myacp{RE} ',
                'Real environments (RE)'         : ' \\myacp{RE} ',
                'real environments (RE)'         : ' \\myacp{RE} ',
                'Real Environments'              : ' \\myacp{RE} ',
                'Real environments'              : ' \\myAcp{RE} ',
                'real environments'              : ' \\myacp{RE} ',
                'Real Environment (RE)'         : ' \\myac{RE} ',
                'Real environment (RE)'         : ' \\myac{RE} ',
                'real environment (RE)'         : ' \\myac{RE} ',
                'Real Environment'              : ' \\myac{RE} ',
                'Real environment'              : ' \\myac{RE} ',
                'real environment'              : ' \\myac{RE} ',

                ' RtD '                         : ' \\myac{RtD} ',
                ' RtD.'                         : ' \\myac{RtD} ',
                ' RtD,'                         : ' \\myac{RtD} ',
                'Research Through Design (RtD)' : ' \\myac{RtD} ',
                'Research through Design (RtD)' : ' \\myac{RtD} ',
                'Research through design (RtD)' : ' \\myac{RtD} ',
                'research through design (RtD)' : ' \\myac{RtD} ',
                'Research-Through-Design (RtD)' : ' \\myac{RtD} ',
                'Research-through-Design (RtD)' : ' \\myac{RtD} ',
                'Research-through-design (RtD)' : ' \\myac{RtD} ',
                'research-through-design (RtD)' : ' \\myac{RtD} ',
                'Research Through Design'       : ' \\myac{RtD} ',
                'Research through Design'       : ' \\myac{RtD} ',
                'Research through design'       : ' \\myac{RtD} ',
                'research through design'       : ' \\myac{RtD} ',
                'Research-Through-Design'       : ' \\myac{RtD} ',
                'Research-through-Design'       : ' \\myac{RtD} ',
                'Research-through-design'       : ' \\myAc{RtD} ',
                'research-through-design'       : ' \\myac{RtD} ',

                ' RfD '                     : ' \\myac{RfD} ',
                ' RfD.'                     : ' \\myac{RfD} ',
                ' RfD,'                     : ' \\myac{RfD} ',
                'Research For Design (RfD)' : ' \\myac{RfD} ',
                'Research for Design (RfD)' : ' \\myac{RfD} ',
                'Research for design (RfD)' : ' \\myac{RfD} ',
                'research for design (RfD)' : ' \\myac{RfD} ',
                'Research-For-Design (RfD)' : ' \\myac{RfD} ',
                'Research-for-Design (RfD)' : ' \\myac{RfD} ',
                'Research-for-design (RfD)' : ' \\myac{RfD} ',
                'research-for-design (RfD)' : ' \\myac{RfD} ',
                'Research For Design'       : ' \\myac{RfD} ',
                'Research for Design'       : ' \\myac{RfD} ',
                'Research for design'       : ' \\myAc{RfD} ',
                'research for design'       : ' \\myac{RfD} ',
                'Research-For-Design'       : ' \\myac{RfD} ',
                'Research-for-Design'       : ' \\myac{RfD} ',
                'Research-for-design'       : ' \\myAc{RfD} ',
                'research-for-design'       : ' \\myac{RfD} ',

                ' RwD '                      : ' \\myac{RwD} ',
                ' RwD.'                      : ' \\myac{RwD} ',
                ' RwD,'                      : ' \\myac{RwD} ',
                'Research With Design (RwD)' : ' \\myac{RwD} ',
                'Research with Design (RwD)' : ' \\myac{RwD} ',
                'Research with design (RwD)' : ' \\myac{RwD} ',
                'research with design (RwD)' : ' \\myac{RwD} ',
                'Research-With-Design (RwD)' : ' \\myac{RwD} ',
                'Research-with-Design (RwD)' : ' \\myac{RwD} ',
                'Research-with-design (RwD)' : ' \\myac{RwD} ',
                'research-with-design (RwD)' : ' \\myac{RwD} ',
                'Research With Design'       : ' \\myac{RwD} ',
                'Research with Design'       : ' \\myac{RwD} ',
                'Research with design'       : ' \\myAc{RwD} ',
                'research with design'       : ' \\myac{RwD} ',
                'Research-With-Design'       : ' \\myac{RwD} ',
                'Research-with-Design'       : ' \\myac{RwD} ',
                'Research-with-design'       : ' \\myAc{RwD} ',
                'research-with-design'       : ' \\myac{RwD} ',

                ' UGC '                         : ' \\myac{UGC} ',
                ' UGC.'                         : ' \\myac{UGC} ',
                ' UGC,'                         : ' \\myac{UGC} ',
                'User-Generated Content (UGC)'  : ' \\myac{UGC} ',
                'User-generated content (UGC)'  : ' \\myac{UGC} ',
                'User-generated content (UGC)'  : ' \\myac{UGC} ',
                'User Generated Content (UGC)'  : ' \\myac{UGC} ',
                'User generated content (UGC)'  : ' \\myac{UGC} ',
                'user generated content (UGC)'  : ' \\myac{UGC} ',
                'User-Generated Content'        : ' \\myac{UGC} ',
                'User-generated content'        : ' \\myac{UGC} ',
                'User-generated content'        : ' \\myac{UGC} ',
                'User Generated Content'        : ' \\myac{UGC} ',
                'User generated content'        : ' \\myac{UGC} ',
                'user generated content'        : ' \\myac{UGC} ',
                '"user-generated content," (UGC)'        : ' \\myac{UGC} ',
                '``user-generated content," (UGC)'        : ' \\myac{UGC} ',
                'user-generated content," (UGC)'        : ' \\myac{UGC} ',

                ' RVC '                         : ' \\myac{RVC} ',
                ' RVC.'                         : ' \\myac{RVC} ',
                ' RVC,'                         : ' \\myac{RVC} ',
                'Reality-Virtuality Continuum (RVC)'  : ' \\myac{RVC} ',
                'Reality-virtuality continuum (RVC)'  : ' \\myac{RVC} ',
                'reality-virtuality continuum (RVC)'  : ' \\myac{RVC} ',
                'Reality Virtuality Continuum (RVC)'  : ' \\myac{RVC} ',
                'Reality virtuality continuum (RVC)'  : ' \\myAc{RVC} ',
                'reality virtuality continuum (RVC)'  : ' \\myac{RVC} ',
                'Reality-Virtuality Continuum'        : ' \\myac{RVC} ',
                'Reality-virtuality continuum'        : ' \\myac{RVC} ',
                'reality-virtuality continuum'        : ' \\myac{RVC} ',
                'Reality Virtuality Continuum'        : ' \\myac{RVC} ',
                'Reality virtuality continuum'        : ' \\myAc{RVC} ',
                'reality virtuality continuum'        : ' \\myac{RVC} ',

                ' UE '                          : ' \\myac{UE} ',
                ' UE.'                          : ' \\myac{UE} ',
                ' UE,'                          : ' \\myac{UE} ',
                'Unreal Engine (UE)'          : ' \\myac{UE} ',
                'Unreal engine (UE)'          : ' \\myac{UE} ',
                'unreal engine (UE)'          : ' \\myac{UE} ',
                'Unreal Engine'               : ' \\myac{UE} ',
                'Unreal engine'               : ' \\myAc{UE} ',
                'unreal engine'               : ' \\myac{UE} ',
                
                ' UTS '                     : ' \\myac{UTS} ',
                ' UTS.'                     : ' \\myac{UTS} ',
                ' UTS,'                     : ' \\myac{UTS} ',
                '"Under the Skin"'          : ' \\myacl{UTS} ',
                '"Under the Skin"'          : ' \\myacl{UTS} ',
                '``Under the Skin"'         : ' \\myacl{UTS} ',
                '\'Under the Skin\''        : ' \\myacl{UTS} ',
                '"Under The Skin"'          : ' \\myacl{UTS} ',
                '"Under The Skin"'          : ' \\myacl{UTS} ',
                '``Under The Skin"'         : ' \\myacl{UTS} ',
                '\'Under The Skin\''        : ' \\myacl{UTS} ',
                'Under The Skin'            : ' \\myacl{UTS} ',
                'Under the Skin'            : ' \\myacl{UTS} ',

                ' UV '                       : ' \\myac{UV} ',
                ' UV.'                       : ' \\myac{UV} ',
                ' UV,'                       : ' \\myac{UV} ',
                'Ultra-Violet (UV)'          : ' \\myac{UV} ',
                'Ultra-violet (UV)'          : ' \\myAc{UV} ',
                'ultra-violet (UV)'          : ' \\myac{UV} ',
                'Ultra Violet (UV)'          : ' \\myac{UV} ',
                'Ultra violet (UV)'          : ' \\myAc{UV} ',
                'ultra violet (UV)'          : ' \\myac{UV} ',
                'Ultraviolet (UV)'          : ' \\myac{UV} ',
                'ultraviolet (UV)'          : ' \\myac{UV} ',
                'Ultra-Violet'               : ' \\myac{UV} ',
                'Ultra-violet'               : ' \\myAc{UV} ',
                'ultra-violet'               : ' \\myac{UV} ',
                'Ultra Violet'               : ' \\myac{UV} ',
                'Ultra violet'               : ' \\myAc{UV} ',
                'ultra violet'               : ' \\myac{UV} ',
                'Ultraviolet'               : ' \\myAc{UV} ',
                'ultraviolet'               : ' \\myac{UV} ',

                ' VE '                      : ' \\myac{VE} ',
                ' VE.'                      : ' \\myac{VE} ',
                ' VE,'                      : ' \\myac{VE} ',
                ' VEs '                      : ' \\myacp{VE} ',
                ' VEs.'                      : ' \\myacp{VE} ',
                ' VEs,'                      : ' \\myacp{VE} ',
                'Virtual Environments (VE)'  : ' \\myAcp{VE} ',
                'Virtual environments (VE)'  : ' \\myAcp{VE} ',
                'virtual environments (VE)'  : ' \\myacp{VE} ',
                'Virtual Environments'       : ' \\myAcp{VE} ',
                'Virtual environments'       : ' \\myAcp{VE} ',
                'virtual environments'       : ' \\myacp{VE} ',
                'Virtual Environment (VE)'  : ' \\myAc{VE} ',
                'Virtual environment (VE)'  : ' \\myAc{VE} ',
                'virtual environment (VE)'  : ' \\myac{VE} ',
                'Virtual Environment'       : ' \\myac{VE} ',
                'Virtual environment'       : ' \\myac{VE} ',
                'virtual environment'       : ' \\myac{VE} ',

                ' VR '                          : ' \\myac{VR} ',
                ' VR.'                          : ' \\myac{VR} ',
                ' VR,'                          : ' \\myac{VR} ',
                'Virtual Reality (VR)'          : ' \\myac{VR} ',
                'Virtual reality (VR)'          : ' \\myac{VR} ',
                'virtual reality (VR)'          : ' \\myac{VR} ',
                'Virtual Reality'               : ' \\myac{VR} ',
                'Virtual reality'               : ' \\myac{VR} ',
                'virtual reality'               : ' \\myac{VR} ',

                ' XR '                          : ' \\myac{XR} ',
                ' XR.'                          : ' \\myac{XR} ',
                ' XR,'                          : ' \\myac{XR} ',
                'Extended Reality (XR)'         : ' \\myac{XR} ',
                'Extended reality (XR)'         : ' \\myac{XR} ',
                'extended reality (XR)'         : ' \\myac{XR} ',
                'eXtended Reality (XR)'         : ' \\myac{XR} ',
                'Extended Reality'              : ' \\myac{XR} ',
                'Extended-Reality'              : ' \\myac{XR} ',
                'extended reality'              : ' \\myac{XR} ',
                'extended-reality'              : ' \\myac{XR} ',

            }

####################################### MAIN FUNCTIONS ################################################################
          
def generate_tex(input_file):
    try:
        #update docx for current chapter number
        doc_file = input_file
        tex_file = input_file + ".tex"
    except Exception as e:
        print(bold_red,"Update chpater failed:", e , reset)
        return False
      
    if not os.path.exists(doc_file):
        print(bold_red + "Missing file path: " + doc_file + reset)
        return False
    else:
        print(italic, "Processing File:", (doc_file.split('\\Chapters\\')[1]), reset)
        conversion_result = convert_docx_to_tex(doc_file, tex_file)
        if conversion_result:
            try:
                fixes_result = manual_fixes(input_file)
                return fixes_result
            except:
                print(bold_red + "Manual Fixes failed" + reset)
                return False
    return False


def convert_docx_to_tex(input_file, output_file):
    try:
        pypandoc.convert_file(input_file, to_format, input_format, outputfile=output_file, extra_args=extra_args)
        return True
    except Exception as e:
        print(bold_red, "Pandoc ERROR:", str(e), reset)
        return False

def manual_fixes(input_file):
    #create and verify zip version of document then unzip to temp
    doc_file = input_file
    tex_file = input_file + ".tex"
    result = False
    captions = []
    image_names = []
    subdoc_captions = []
    had_error = False
    pattern = r'^\\includegraphics\[width=.*in,height=.*in\]{'
    try:
        image_names = get_original_image_names(doc_file)    
    except Exception as e:
        print(bold_red, "Failed to find image names:", e, reset)
        had_error = True

    try: #Copy each caption for figures into their own temporary docx so they can be individually pandoc converted and crorrefs can be linked
        subdoc_captions = extract_paragraphs_by_style(doc_file, "Caption")
    except Exception as e:
        print(bold_red, "Failed to find captions by style:", e, reset)
        had_error = True

    try:
        crossrefs = find_text_with_style(doc_file, "CrossReference")
    except Exception as e:
        print(bold_red, "Failed to find crossreferences by style:", e, reset)
        had_error = True

    try: #remove empty captions          
        for c in captions:
            if c == "":
                print(yellow, "Warning: Removed empty caption", reset)
                had_error = True
                captions.remove(c)
    except Exception as e:
        print(bold_red, "Remove empty captions failed:", e, reset)
        had_error = True

    #Console line to inform
    print(green, "Number of Images = ", len(image_names))
    print(green, "Number of Captions = ", len(subdoc_captions), reset)
    if len(image_names) != len(subdoc_captions):
        print(bold_red, "ERROR: Figure MISMATCH", reset)
        had_error = True
    for name in image_names:
        i = image_names.index(name)
        print(yellow, f"Progress: {i}/{len(image_names)}", reset, end='\r')
        delimiter = "Note."
        before_note = ""
        #search_pattern = before_note.encode('utf-8', 'escape')
        after_note = ""
        end_search = ""
        saved_note = ""
        has_note = False
        latex_caption = ""
        cap_search = ""
        search_pattern = b""

        try : #Remove old end figure command
            replace_line_with_pattern(tex_file, r"\\end{fignos:no-prefix-figure-caption}", "")
            result = True
        except:
            print(bold_red, "Error changing to \\end{figure}", reset)
            had_error = True

        try:
            name = os.path.basename(name)
        except Exception as e:
            print(bold_red, "ERROR: Failed to get OS path for Image: ", name, "   Error is:", e, reset)
            had_error = True

        try:
            replace_line_with_pattern(tex_file, r"\\begin{fignos:no-prefix-figure-caption}", r"\begin{figure}")
            result = True
        except Exception as e:
            print(bold_red, "Error changing \\begin{figure} WITH ERROR:", e, reset)
            had_error = True

        label_name, *_ = name.split(".")
        label_name = r"    \label{fig:" + label_name + "}"
        replacement = "    \\includegraphics[width=\\textwidth]{" + "Figures/" + name + "}"
        try:#Make figures text width and add image filepath
            replace_line_with_pattern(tex_file, pattern, replacement)
            result = True    
        except Exception as e:
            print(bold_red, "Error adjusting \\includegraphics: ", e, reset)
            had_error = True
        try:
            modified_text, unedited_line = replace_youtube_link_with_command(tex_file, 'Youtube')
            if modified_text is not None:
                add_new_line_of_text_above_word(tex_file, unedited_line.decode('utf-8', 'replace'), modified_text.decode('utf-8', 'replace')[:-1])
        except Exception as e:
            print(bold_red, e, reset)
        try:
            modified_text, unedited_line = replace_youtube_link_with_command(tex_file, 'Vimeo')
            if modified_text is not None:
                add_new_line_of_text_above_word(tex_file, unedited_line.decode('utf-8', 'replace'), modified_text.decode('utf-8', 'replace')[:-1])
        except Exception as e:
            print(bold_red, e, reset)
        try:          

            latex_caption = pypandoc.convert_file(subdoc_captions[i], to_format, input_format, extra_args=extra_args)
            latex_caption = replace_caption_crossreferences(latex_caption)
            if delimiter in latex_caption:
                before_note, after_note = latex_caption.split(delimiter, 1)
                before_note = r"    \caption{" + before_note + "}"

                after_note = r"\textit{" + delimiter + " }" + after_note.strip() # Include the delimiter in the "after_note" string
                after_note = r"" + after_note
                
                after_note = "    " + r"    \raggedright{\small{" + after_note + "}}"
                #search_pattern = get_first_4_words(latex_caption)
                search_pattern = truncate_and_encode(latex_caption, 50)
                saved_note = replace_line_with_pattern(tex_file, search_pattern, after_note, True) #Replace the original under-image caption with only the "after-note" section
                has_note = True
            else:
                before_note = r"" + latex_caption
                before_note = before_note.strip()
                before_note = r"    \caption{" + before_note + "}"
                caption_replacement = ""
                #search_pattern = get_first_4_words(latex_caption)
                search_pattern = truncate_and_encode(latex_caption, 50)
                replace_line_with_pattern(tex_file, search_pattern, caption_replacement) #Delete original caption that was under the image
                has_note = False

            add_new_line_of_text_above_word(tex_file, replacement, before_note) #Add the figure title/short caption above image
            result = True
        except Exception as e:
            print(bold_red, "Failed to correct caption:", e, reset)
            had_error = True
        try:      
            cap_search =r"\\caption{" + search_pattern.decode("utf-8", 'ignore')  #Ignore unbound 
            #time.sleep(0.1)
            #add_line_below_pattern(tex_file, cap_search, label_name) #Add a label for cross-referencing
            add_new_line_of_text_above_word(tex_file, replacement, label_name)
            
            result = True
        except Exception as e:
            print(bold_red, "Error adding label:", e, reset)
            had_error = True
        try:
            if cap_search != "":
                #add_line_below_pattern(tex_file, cap_search, r"    \centering") #Centre the image
                add_new_line_of_text_above_word(tex_file, label_name, r"    \centering")
        except Exception as e:
            print(bold_red, "Error adding centering cmd:", e, reset)
        try: 
            if has_note:
                end_search =  saved_note
            else:
                end_search = replacement 
            add_new_line_of_text_below_word(tex_file, end_search, r"\end{figure}") #Close the figure environment
            result = True
        except Exception as e:
            print(bold_red, "Error adding \\end{Figure} line:", e, reset)
            had_error = True
    
    try:
        acrocount = 0
        numreplaced = 0
        for acro in acronyms:
            numreplaced = replace_string_in_tex(tex_file, acro, acronyms.get(acro))
            acrocount = acrocount + numreplaced
        print(green,'Replaced ', acrocount, ' acronyms in this chapter', reset)
    except :
        print(bold_red, "ERROR: Acronyms failed", reset)        
    try:
        replace_crossreferences(tex_file) #Run the function that replaces cross refs based on style with LATEX version
    except Exception as e:
        print(bold_red, "Replace captions failed:", e, reset)
        had_error = True

    if len(image_names) == 0 and had_error == False:
        return True
    return result

###################################### SHORT FUNCTIONS ###################################################################

def sync_chapter_to_overleaf(project_ref, chapter_num):                  
    new_tex = update_file_number(chapter_num, ".tex")
    old_tex = "Chapter" + chapter_num + ".tex"
    project_ref.update_file(old_tex, new_tex)  


def update_file_number(chapter_number, extension):
    file = os.path.join(doc_path, 'Chapter')
    file = file + str(chapter_number) + extension
    return file

def get_original_image_names(docx_file):
    doc = Document(docx_file)
    image_names = []
    refs = []
    ordered_images = []
    all_rids = []

    try:
        for rel in doc.part.rels.values():
            if "image" in rel.reltype:
                try:
                    image_name = rel._target.split("/")[-1]
                    image_name = r"" + image_name
                    image_names.append(image_name)
                    ref = rel.rId 
                    #print(ref)
                    refs.append(ref)
                except:
                    pass
    except Exception as e:
        print(bold_red, "rel.values failed:", e, reset)
    try:
        all_rids = get_rid_order(docx_file)
    except Exception as e:
        print(bold_red, "All rID's failed:",e, reset)
    try:
        for rid in all_rids:
            if rid in refs:
                pass
            else:
                all_rids.remove(rid)
    except Exception as e:
        print(bold_red, "Fix all rID list fialed:",e, reset)

    #reordered_rids = sorted(refs, key=lambda x: all_rids.index(x))
    try:
        ordered_images = [c for _, c in sorted(zip(refs, image_names), key=lambda x: all_rids.index(x[0]))]
    except Exception as e:
        print(bold_red, "order_images Failed:", e, reset)
    return ordered_images

def remove_trailing_whitespace(string):
    return string.rstrip()

def get_first_4_words(string):
    words = string.split()
    first_4_words = words[:4]
    return ' '.join(first_4_words)

def get_last_4_words(string):
    words = string.split()
    last_4_words = words[-4:]
    return ' '.join(last_4_words)

def truncate_and_encode(input_string, max_length=30):
    # Check if the input string length is less than or equal to max_length
    if len(input_string) <= max_length:
        # Encode the entire string to UTF-8 with escape characters
        return input_string.encode('utf-8', 'escape')
    else:
        # Truncate the string to the first 30 characters and encode to UTF-8 with escape characters
        return input_string[:max_length].encode('utf-8', 'escape')

def add_line_above_first_line(tex_file, new_line):
    with open(tex_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    lines.insert(0, new_line + '\n')

    with open(tex_file, 'w', encoding='utf-8') as file:
        file.writelines(lines)

def find_replace_unknown(original, replacement):
    # Escape special characters in the original string
    escaped_original = re.escape(original)
    
    # Create a regular expression pattern with the escaped original string
    pattern = re.compile(escaped_original)
    
    # Replace the unknown characters in the replacement string using the pattern
    modified = re.sub(pattern, replacement, original)
    
    return modified

def replace_crossreferences(file_path):
    pattern = r'\\emph{\\textbf{\\ul{([^{}]+)}}}'

    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    def repl(match):
        cross_ref = match.group(1)
        cross_ref = cross_ref.replace("\\", "")
        return r'\ref{' + cross_ref + '}'

    modified_content = re.sub(pattern, repl, content)

    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(modified_content)

def replace_first_occurrence(file_path, word_to_replace, replacement):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    # Find the index of the first occurrence of the word
    index = content.find(word_to_replace)

    if index != -1:
        # Replace the word with the desired replacement
        updated_content = content[:index] + replacement + content[index + len(word_to_replace):]

        # Write the updated content back to the file
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(updated_content)

def replace_line_with_pattern(file_path, pattern, replacement, is_after_caption=False):
    replaced = False  # Flag to track if replacement has been made
    delimiter = 'Note.'
    before_note = ''
    after_note = ''

    # Check if 'pattern' is of type bytes
    if isinstance(pattern, bytes):
        pattern = pattern.decode('utf-8')

    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    with open(file_path, 'w', encoding='utf-8') as file:
        for line in lines:
            if not replaced:
                if pattern in line or re.search(pattern, line):
                    if is_after_caption:
                        before_note, after_note = line.split(delimiter, 1)
                        after_note = r"\textit{" + delimiter + " }" + after_note.strip()  # Include the delimiter in the "after_note" string
                        after_note = r"" + after_note
                        after_note = r"    \raggedright{\small{" + replace_caption_crossreferences(after_note) + "}}"
                        file.write(after_note + '\n')
                    else:
                        file.write(replacement + '\n')
                    replaced = True  # Set the flag to True after making the replacement
                else:
                    file.write(line)
            else:
                file.write(line)
    return after_note

def add_line_above_pattern(file_path, pattern, new_line):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    modified_content = re.sub(pattern, new_line + r'\n\g<0>', content)

    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(modified_content)

def add_line_below_pattern(file_path, pattern, new_line):
    with open(file_path, 'rb') as file:
        content = file.read()

    # Ensure that the pattern is a bytes object
    if isinstance(pattern, str):
        pattern = pattern.encode('utf-8')

    # Use bytes pattern to search in content
    pattern = re.compile(pattern)
    modified_content = pattern.sub(lambda x: x.group(0) + b'\n' + new_line.encode('utf-8'), content, count=1)

    with open(file_path, 'wb') as file:
        file.write(modified_content)



def find_text_with_style(document_path, style_name):
    doc = Document(document_path)
    text_with_style = []

    for paragraph in doc.paragraphs:
        if paragraph.style.name == style_name:
            text_with_style.append(paragraph.text)

    return text_with_style

def find_and_replace(file_path, search_text, replace_text):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    try:
        modified_content = content.replace(search_text, replace_text)
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(modified_content)
    except Exception as e:
        print(bold_red, "Find and Replace Failed:", e, reset)

def get_rid_order(docx_file):
    # Open the document using python-docx
    doc = Document(docx_file)

    # Retrieve the document.xml part
    document_part = doc.part.document

    # Convert the document XML blob to string
   # xml_string = document_part.part.blob.decode('utf-8')
    root = None
    drawing_elements = []
    namespaces = {}
    try:
        # Parse the document XML using lxml
        root = etree.fromstring(doc.part.blob, parser=None)
    except Exception as e:
        print(bold_red, "XML Parsing Failed:", e, reset)

    try:
    # Register the 'w' namespace prefix
        namespaces = {
            'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
            'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
            'cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
            'cx3': 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
            'cx4': 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
            'cx5': 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
            'cx6': 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
            'cx7': 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
            'cx8': 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'aink': 'http://schemas.microsoft.com/office/drawing/2016/ink',
            'am3d': 'http://schemas.micro...17/model3d',
            'o': 'urn:schemas-microsoft-com:office:office',
            'oel': 'http://schemas.microsoft.com/office/2019/extlst',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'v': 'urn:schemas-microsoft-com:vml',
            'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'w10': 'urn:schemas-microsoft-com:office:word',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
            'w16cex': 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
            'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
            'w16': 'http://schemas.microsoft.com/office/word/2018/wordml',
            'w16du': 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
            'w16sdtdh': 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
            'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
            'wpg': 'http://schemas.microsoft.com/office/' }

            
        # Find all w:drawing elements in the document.xml
        #drawing_elements = root.findall(".//w:drawing", namespaces=namespaces)
        if root != None:      
            drawing_elements = root.findall(".//a:blip", namespaces=namespaces)
        else:
            print(bold_red, "Get RID Order Failed: doc.part.blob.root is None", reset)

    except Exception as e:
        print(bold_red, "w:drawing search failed:", e, reset)

    # Retrieve the rId values in the order they appear in the document.xml
    rid_order = []
    try:
        for drawing_element in drawing_elements:
            #blip_element = drawing_element.find(".//a:blip", namespaces=namespaces)
            if drawing_element is not None:
                embed_attrib_name = QName(namespaces['r'], 'link')
                try:
                    str_attrib = str(embed_attrib_name)
                    r_id = drawing_element.attrib[str_attrib]
                    rid_order.append(r_id)
                except KeyError:
                    print(bold_red,"Failed to get image link:", KeyError, reset)
    except Exception as e:
        print(bold_red, "rid list failed with a non-key-error:", e)
    return rid_order


def add_new_line_of_text_above_word(tex_file, specific_word, new_line_text):
    # Read the content of the tex file
    with open(tex_file, 'r', encoding='utf-8') as file:
        content = file.readlines()

    # Iterate over the lines and add a new line of text above the first line with the specific word
    modified_content = []
    word_encountered = False  # Flag variable to track if the specific word has been encountered
    for line in content:
        if specific_word in line and not word_encountered:
            modified_content.append(new_line_text + '\n')  # Add the new line of text
            word_encountered = True  # Set the flag to True after encountering the word
        modified_content.append(line)

    # Write the modified content back to the tex file
    with open(tex_file, 'w', encoding='utf-8') as file:
        file.writelines(modified_content)

def add_new_line_of_text_below_word(tex_file, specific_word, new_line_text):
    # Read the content of the tex file
    with open(tex_file, 'r', encoding='utf-8') as file:
        content = file.readlines()

    # Iterate over the lines and add a new line of text above the first line with the specific word
    modified_content = []
    word_encountered = False  # Flag variable to track if the specific word has been encountered
    for line in content:
        modified_content.append(line)
        if specific_word in line and not word_encountered:
            modified_content.append(new_line_text + '\n')  # Add the new line of text
            word_encountered = True  # Set the flag to True after encountering the word

    # Write the modified content back to the tex file
    with open(tex_file, 'w', encoding='utf-8') as file:
        file.writelines(modified_content)

def extract_special_words(input_string):
    special_words = re.findall(r'\S*(?:fig:|_|\-)\S*', input_string)
    return special_words

def extract_paragraphs_by_style(docx_path, style_name):
    doc = Document(docx_path)
    paragraphs_with_style = []
    temp_files = []

    for i, paragraph in enumerate(doc.paragraphs):
        if paragraph.style.name == style_name:
            new_doc = Document()
            new_paragraph = new_doc.add_paragraph()

            # Copy run-level formatting from the original paragraph to the new paragraph
            for run in paragraph.runs:
                new_run = new_paragraph.add_run(run.text)
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline
                # ... copy other font properties as needed

            temp_file = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
            temp_files.append(temp_file.name)
            new_doc.save(temp_file.name)

            paragraphs_with_style.append(temp_file.name)

    return paragraphs_with_style

def replace_caption_crossreferences(content):
    pattern = r'\\emph{\\textbf{\\ul{([^{}]+)}}}'

    def repl(match):
        cross_ref = match.group(1)
        cross_ref = cross_ref.replace("\\", "")
        return r'\ref{' + cross_ref + '}'

    modified_content = re.sub(pattern, repl, content)
    return modified_content




def handle_failed_chapter(e):
    print(bold_red, "ERROR: Failed to convert Chapter", " Error:", e, reset)


def replace_youtube_link_with_command(tex_file_path, website):
    try:
        # Read the content of the .tex file
        with open(tex_file_path, 'r', encoding='utf-8') as file:
            tex_content = file.read()
        
        if website == 'Youtube':
            # Define regular expressions for the YouTube links
            regex = r'https://youtu\.be/[^\s]+|https://www\.youtube\.com/[^\s]+'
        elif website == 'Vimeo':
            # Define regular expressions for the Vimeo links
            regex = r'https://vimeo\.com/[^\s]+'
        else:
            regex = r'https://youtu\.be/[^\s]+|https://www\.youtube\.com/[^\s]+'
            
        # Search for the first YouTube link in the file
        match = re.search(regex, tex_content)

        if match:
            # Extract the matched YouTube link
            website_link = match.group(0)
            shortlink = re.sub(r'https://youtu\.be/', '', website_link)
            shortlink = re.sub(r'https://youtube\.com/', '', shortlink)
            shortlink = re.sub(r'https://vimeo\.com/', '', shortlink)

            if website == 'Youtube':
                # Format the YouTube link
                formatted_website_link = r'\youtube{' + shortlink + '}'
            else:
                formatted_website_link = r'\vimeo{' + shortlink + '}'
        else:
            return None, None
        
        # Find the entire unedited line where the URL was found
        lines = tex_content.split('\n')
        for i, line in enumerate(lines):
            if re.search(regex, line):
                unedited_line = lines[i]
                break
        else:
            unedited_line = None

        # Encode the modified text and unedited line into UTF-8 bytes with ASCII escape characters
        modified_bytes = formatted_website_link.encode('utf-8')
        unedited_bytes = unedited_line.encode('utf-8') if unedited_line is not None else None

        #return formatted_youtube_link, unedited_line
        return modified_bytes, unedited_bytes
    except Exception as e:
        return None, None


def replace_string_in_tex(tex_file, search_string, replace_string):
    # Define a custom replacement function
    def custom_replace(match):
        return replace_string

    # Read the content of the .tex file
    with open(tex_file, 'r', encoding='utf-8') as file:
        tex_content = file.read()

    # Perform the replacement using the custom function
    modified_content, replacements_count = re.subn(re.escape(search_string), custom_replace, tex_content, flags=re.MULTILINE)

    # Write the modified content back to the .tex file
    with open(tex_file, 'w', encoding='utf-8') as file:
        file.write(modified_content)

    return replacements_count


def get_docx_files(directory):
    docx_files = []
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".docx"):
                file_path = os.path.join(root, file)
                docx_files.append(file_path)
    
    return docx_files

#***************** RUN FULL CONVERSION *********************************
#Pre convert locally for all files in current folder
docx_files = get_docx_files(current_directory)

for d in docx_files:
    print(d)
    try:
        gen_result = generate_tex(d)
        print("Finished Converting file:", d)
    except Exception as e:
        print("ERROR: Failed to convert", e)
print("Finished converting all files")



