# goal:  match two rows, then replace the id of messtedup table row with previous table row
import pandas as panda

# Grab only the path and oid from the correct sheet
correct_sheet = panda.read_excel('../Original.xlsx', usecols=['path', '_id.$oid'])
# broken_sheet = panda.read_excel('../Messedup.xlsx', usecols=['path', '_id.$oid'] ).head()
broken_sheet = panda.read_excel('../Cleanup_FIXED.xlsx')

is_match = lambda path : lambda row : row['path'] == path

def operate_on_match(operation, path_to_match):
  operation(correct_sheet.loc[is_match(path_to_match)])

def replace_oid(index):
  def replace(matching_row):
    for i, oid in matching_row['_id.$oid'].items():
      broken_sheet.loc[index, ['_id.$oid']] = oid
  return replace

def fix():
  for index, path in broken_sheet['path'].items():
    operate_on_match(replace_oid(index), path)

fix()
broken_sheet.to_excel('Fixed.xlsx')
