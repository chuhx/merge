# This module merge the info of RCW into register table.
#
# First version was written by Larry Chu.
# $Author: lchu $
# $Revision: 1.20 $
# $Date: 2015/10/20 11:30:00 $

import ComExcel
import os
import copy
import re
import shutil


def extractRcwFrom(fname):
	'''Extract RCW info from an xls file'''
	print '*'*70, '\nStart extract info from %s\n'%fname, '*'*70
	colStart = 2
	rowScope = range(1, 300)
	rcwInfo = {}
	for sht in ['F0 RC0', 'F0 8-bit RCW',  'F1 RC0', 'F1 8-bit RCW', 'F4 4-bit RCW', 'F4 8-bit RCW', 'F7 8-bit RCW']:
		f = ComExcel.ExcelComObj(sheetnum=sht, \
				filename=os.path.join(os.getcwd(),fname))
		print sht
		for row in rowScope:
			cellText = f.getCellText(row, colStart).strip()
			if not (cellText.startswith('RC') or re.match('^F\dRC\w\w', cellText) ):
				continue
			rcwId = cellText.split(':')[0].strip().upper()
			print rcwId + '-'*40
			if re.search('RC\wX$', rcwId): colWidth = 13 # 8-bit RCW
			elif re.search('RC0\w$', rcwId): colWidth = 9 # 4-bit RCW
			else: continue
			rcwInfo[rcwId] = []
			for rowOfs in range(3,50):
				#if f.borderChk(row+rowOfs, colStart, row+rowOfs, colStart) != 1:
					#break
				if f.getCellText(row+rowOfs, colStart) not in ['x','1','0',]:
					break
				bitLenStr = f.getCellText(row+rowOfs, colStart+colWidth-4).strip()
				if bitLenStr != u'':
					attr = f.getCellText(row+rowOfs, colStart+colWidth-3)
					regDef = f.getCellText(row+rowOfs, colStart+colWidth-2)
					regName = f.getCellText(row+rowOfs, colStart+colWidth-5).strip()
					comment = 'Name: %s\nAttribute: %s\nDescription: %s\n' \
							%(regName, attr, regDef)
					if regName in ['NV_MPR0', 'NV_MPR1', 'NV_MPR2', 'LCOM_VREF']:
						if   regName == 'NV_MPR0': defaultVal = 0x55
						elif regName == 'NV_MPR1': defaultVal = 0x33
						elif regName == 'NV_MPR2': defaultVal = 0x0f
						elif regName == 'LCOM_VREF': defaultVal = 0x5
						comment += "Default: 'h%02x\n"%defaultVal
					bitLen = int(bitLenStr)
					aReg = {'name': copy.copy(regName), 
							'bit_len': copy.copy(bitLen), 
							'comment': copy.copy(comment), 
							}
					rcwInfo[rcwId].append(aReg)

		# f.close()
	return rcwInfo

def merge(rcwInfo, fname):
	''' Merge RCW info into register table file'''
	print '\n\n' + '*'*70, '\nStart replace RCWs in %s\n'%fname, '*'*70
	for sht in ['Function0', 'Function1', 'Function4', 'Function7']:
		f = ComExcel.ExcelComObj(sheetnum=sht, filename=os.path.join(os.getcwd(),fname))
		for row in range(5, 17):
			for col in range(2, 34):
				if f.getCellText(row, col) in rcwInfo.keys():
					rcwId = f.getCellText(row, col)
					if rcwInfo[rcwId] == []: continue
					print '%s found'%rcwId
					if re.search('RC\wX$', rcwId): rcwBitLen = 8
					elif re.search('RC\w\w$', rcwId): rcwBitLen = 4
					else: raise Exception('Invalid RCW: %s'%rcwId)
					f.unmergeCell(row, col)
					accum = 0
					for reg in rcwInfo[rcwId]:
						accum += reg['bit_len']
						regCol =  col + rcwBitLen - accum
						f.setCell(row, regCol, reg['name'] )
						f.addComment(row, regCol, reg['comment'] )
						f.setCommentFontBoldOff(row, regCol)
						f.setCommentRectangle(row, regCol, 225, 150)
						f.colMerge(row, regCol, row, regCol + reg['bit_len'] - 1 )
		f.save()


def findLatestFile(head):
	''' Get the name of the latest file within a serie of versions 
	whose name all start with argument head. '''
	files = []
	for filename in os.listdir(os.getcwd()):
		if filename.startswith(head):
			files.append(filename)
	if len(files):
		return os.path.join(os.getcwd(), max(files))
	else:
		raise Exception, 'File that starts with %s not found'%head


def run():
	print '*'*70 + '\n' + "<Copy of cb_register_table.xls> to <cb_register_table.xls>"
	f = ComExcel.ExcelComObj(filename = os.path.join(os.getcwd(), 'cb_register_table.xls'))
	f.close()
	os.remove( 'cb_register_table.xls')
	shutil.copy('Copy of cb_register_table.xls', 'cb_register_table.xls')
	merge(extractRcwFrom(findLatestFile('crater_RB_CB_control_word')), \
			findLatestFile('cb_register_table') )

if __name__ == '__main__':
	try: 
		run()
	except: 
		import traceback; traceback.print_exc(); 
	finally:
		raw_input("\n---Press ENTER to quit---")

