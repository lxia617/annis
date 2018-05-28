from xml.etree import ElementTree
from collections import defaultdict
import pprint
import xlsxwriter

input_file='/Users/lixia/Documents/ANNIS/pepper/eaf/AL_RM.eaf'
out_file='/Users/lixia/Documents/ANNIS/pepper/eaf/out/out.xlsx'
with open(input_file, 'rt') as f:
    tree = ElementTree.parse(f)

# time_slot = {}
# for node in tree.iter('TIME_SLOT'):
#     slot = node.attrib.get('TIME_SLOT_ID')
#     value = node.attrib.get('TIME_VALUE')
#     if slot and value:
#         time_slot[slot] = value
#         print(slot + ":"+ value)
#     else:
#         print('error')

align_ids = defaultdict(list)
for node in tree.iter('ALIGNABLE_ANNOTATION'):
    ANNOTATION_ID = node.attrib.get('ANNOTATION_ID')
    TIME_SLOT_REF1 = node.attrib.get('TIME_SLOT_REF1')
    TIME_SLOT_REF2 = node.attrib.get('TIME_SLOT_REF2')
    text = node.find('ANNOTATION_VALUE').text
    if text:
        align_ids[ANNOTATION_ID].append(text)
        # print('ANNOTATION_ID:' + ANNOTATION_ID)
        # print('TIME_SLOT_REF1:' + TIME_SLOT_REF1)
        # print('TIME_SLOT_REF2:' + TIME_SLOT_REF2)
        # print('text:' + text)
    else:
        print('error')

ref2id = defaultdict(list)
id2ref={}
id2text={}
for node in tree.iter('REF_ANNOTATION'):
    a_id = node.attrib.get('ANNOTATION_ID')
    ref = node.attrib.get('ANNOTATION_REF')
    text = node.find('ANNOTATION_VALUE').text
    ref2id[ref].append(a_id)
    id2ref[a_id] = ref
    if text:
        id2text[a_id] = text
    else:
        id2text[a_id] = 'none'
        # print('ANNOTATION_ID:' + ANNOTATION_ID)
        # print('ANNOTATION_REF:' + ANNOTATION_REF)
        # print('ANNOTATION_VALUE:' + ANNOTATION_VALUE.text)
        # print('text:' + text)

#utterance_id
annotationValue2AlignId = {}
for node in tree.iter('TIER'):
    if (node.attrib.get('TIER_ID') == 'utterance_id'):
        for child in node.iter('ALIGNABLE_ANNOTATION'):
            annotationValue2AlignId[child.find('ANNOTATION_VALUE').text]=child.attrib.get('ANNOTATION_ID')

alignid2utterance = {}
for node in tree.iter('TIER'):
    if (node.attrib.get('TIER_ID') == 'utterance'):
        for child in node.iter('REF_ANNOTATION'):
            alignid2utterance[child.attrib.get('ANNOTATION_REF')]=child.attrib.get('ANNOTATION_ID')

alignid2translation = {}
for node in tree.iter('TIER'):
    if (node.attrib.get('TIER_ID') == 'utterance_translation'):
        for child in node.iter('REF_ANNOTATION'):
            alignid2translation[child.attrib.get('ANNOTATION_REF')]=child.attrib.get('ANNOTATION_ID')

utterance2utteranceWords = defaultdict(list)
for node in tree.iter('TIER'):
    if (node.attrib.get('TIER_ID') == 'grammatical_words'):
        for child in node.iter('REF_ANNOTATION'):
            utterance2utteranceWords[child.attrib.get('ANNOTATION_REF')].append(child.attrib.get('ANNOTATION_ID'))

#graid: no idea what is it. Do nothing.

# gloss
utteranceWord2enWords = {}
for node in tree.iter('TIER'):
    if (node.attrib.get('TIER_ID') == 'gloss'):
        for child in node.iter('REF_ANNOTATION'):
            utteranceWord2enWords[child.attrib.get('ANNOTATION_REF')]=child.attrib.get('ANNOTATION_ID')


#write to excel.
workbook = xlsxwriter.Workbook(out_file)
corpusSheet = workbook.add_worksheet('corpusSheet')  # corpusSheet.

tok_colum = 1
index_colum = 0
translation_colum = 2
utterance_word_colum = 3
translation_word_colum = 4

corpusSheet.write(0,tok_colum, 'tok')
corpusSheet.write(0,index_colum, 'time')
corpusSheet.write(0,translation_colum, 'translation[tok]')
corpusSheet.write(0,utterance_word_colum, 'utterance_word')
corpusSheet.write(0,translation_word_colum, 'translation[utterance_word]')

row = 1

#merge_range(first_row, first_col, last_row, last_col, data[, cell_format])
for annotationValue in annotationValue2AlignId:
    align_id = annotationValue2AlignId[annotationValue]
    start_row = row
    utterance_id = alignid2utterance[align_id]
    for i in range(len(utterance2utteranceWords[utterance_id])):
        corpusSheet.write_string(start_row + i, utterance_word_colum, id2text[utterance2utteranceWords[utterance_id][i]])
        corpusSheet.write_string(start_row + i, translation_word_colum, id2text[utteranceWord2enWords[utterance2utteranceWords[utterance_id][i]]])

    # for i in range(len(utterance2enWords[utterance_id])):
    #     corpusSheet.write_string(start_row + i, translation_word_colum, id2text[utterance2enWords[utterance_id][i]])

    
    # write multiple columns
    cross_row = len(utterance2utteranceWords[utterance_id])
    if cross_row !=1:
        corpusSheet.merge_range(start_row, index_colum, start_row +cross_row -1, index_colum, annotationValue)
        corpusSheet.merge_range(start_row, tok_colum, start_row +cross_row - 1, tok_colum, id2text[utterance_id])
    else:
        corpusSheet.write_string(start_row, index_colum, annotationValue)
        corpusSheet.write_string(start_row, tok_colum, id2text[utterance_id])

    if ((align_id in alignid2translation) and alignid2translation[align_id] in id2text):
        corpusSheet.merge_range(start_row, translation_colum, start_row +cross_row -1, translation_colum, id2text[alignid2translation[align_id]])
        #corpusSheet.write(i, 2, id2text[alignid2translation[align_id]])
    row = start_row + cross_row 


#metaSheet = workbook.add_worksheet('metaSheet')  # metaSheet.

workbook.close()
