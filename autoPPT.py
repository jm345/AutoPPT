"""
Jordan Marsh

Automates presentations using data gathered from web scraping.

TODO

Short Term
Generate shadows underneath price tag.
Save information from each web crawl to files in a local directory.

Long Term
Migrate from console to GUI for Desktop and Mobile. Use Xamarin for Android Port.

"""

from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from gui import MainWindow
import csv
import operator
import time
import re
import datetime
from tkinter import messagebox,Button,Label,Text,Tk,Entry,Frame,PhotoImage,filedialog,StringVar,Menu,Toplevel,Checkbutton,IntVar,LabelFrame,ttk

PRESENTATION_ID = '1VHoYQIoMOkqrcC02xVi7Np5ttpDHqqIhceTxx61PWGY'		
POINTS_IN_INCH = 72
MARGIN = 0.55
STEELFORT_BASE_URL = 'https://www.steelfort.co.nz'

SCOPES = 'https://www.googleapis.com/auth/presentations'
store = file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
	flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
	creds = tools.run_flow(flow, store)
service = build('slides', 'v1', http=creds.authorize(Http()))

def get_formatted_current_time():

	current_time = datetime.datetime.now()
	year = current_time.year
	month = current_time.strftime('%B')
	day = current_time.strftime('%d')
	return ('%s %s %s' % (day,month,year))

def parse_csv():
	
	specs,titles,prices,images,categories = [],[],[],[],[]
	
	# Open .csv file. Sort the data using the sub categories.
	with open('csv\\big_save5-8-2018.csv', newline='',encoding='utf8') as file:
	
		reader = csv.DictReader(file)
		reader = sorted(reader,key=operator.itemgetter('sub_categories'))
		
		# Append all the required data into their subsequent arrays.
		for row in reader:
				
			category = row['sub_categories']
			spec = row['productSpecs']
			image = row['productImage-src']
			title = row['productTitle']
			price = row['productPrice']
			#spec = row['productSpecs2']
	
			categories.append(category)
			images.append(image)
			titles.append(title)

			spec = re.sub(r'[\n\t\r ]{2,}','\n',spec)
			specs.append(spec)
			
			# If there are multiple price values, use the right most value.
			if price.count('$') > 1:
				price = price[price.rindex('$'):]
			price = price.replace('$','').replace(',','')
			try:
				price = round(float(price) / MARGIN,2)
				prices.append(str(price))
			except:
				#price = re.sub(r'[^\d|\.]','',price)
				#price = round(float(price) / 0.55,2)
				prices.append(str(price))

	return titles,specs,prices,images,categories

	
def create_presentation(titles,specs,prices,images,categories):

	root = Tk()
	main_gui = MainWindow(root)
	i = 0
	slide_position_index = 0
	min_boundary = min(len(prices),len(titles),len(specs),len(images))
	current_date = get_formatted_current_time()
	requests = [
			{
				'createSlide': {
					'objectId': 'TitleSlide',
					'insertionIndex': slide_position_index,
				}
			},
			{
				
				# Create the textbox.
					
					'createShape': {
						'objectId':'TitleTextBox',
						'elementProperties': {
							'pageObjectId': 'TitleSlide',
							'size': {
								'width': {
									'magnitude':9.32 * POINTS_IN_INCH,
									'unit':'PT'
								},
								'height': {
									'magnitude':0.93 * POINTS_IN_INCH,
									'unit':'PT'
								}
							},
							"transform": {
								"scaleX": 1,
								"scaleY": 1,
								"translateX": 0.34 * POINTS_IN_INCH,
								"translateY": 1.96 * POINTS_IN_INCH,
								"unit": "PT"
							}
						},
						'shapeType':'TEXT_BOX',
					}
				},
				{
				
				# Insert the category name from the file into the text box. 
				
						"insertText": {
							'objectId': 'TitleTextBox',
							'text':main_gui.user_title_string,
							'insertionIndex':0
						}
				},
				{
				
				# Update text formatting.
				
					'updateTextStyle': {
						'objectId':'TitleTextBox',
						'fields': 'foregroundColor,fontFamily,fontSize',
						'style': {
							'foregroundColor': {
								'opaqueColor': {
									'rgbColor': {
										'red': 0.27,
										'green': 0.27,
										'blue': 0.27
									}
								}
							},
							'fontFamily': 'Helvetica Neue',
							'fontSize': {
								'magnitude': 48,
								'unit':'PT'
							}
						}
					}
				},
				{
				
				# Update text alignment formatting (Category Slide)
				
					'updateParagraphStyle': {
						'objectId':'TitleTextBox',
						'fields': 'alignment',
						'style': {
							'alignment':'END'
						}
					}
				},
				{
				
				# Create the subtitle text box.
					
					'createShape': {
						'objectId':'SubTitleTextBox',
						'elementProperties': {
							'pageObjectId': 'TitleSlide',
							'size': {
								'width': {
									'magnitude':9.32 * POINTS_IN_INCH,
									'unit':'PT'
								},
								'height': {
									'magnitude':0.67 * POINTS_IN_INCH,
									'unit':'PT'
								}
							},
							"transform": {
								"scaleX": 1,
								"scaleY": 1,
								"translateX": 0.34 * POINTS_IN_INCH,
								"translateY": 2.85 * POINTS_IN_INCH,
								"unit": "PT"
							}
						},
						'shapeType':'TEXT_BOX',
					}
				},
				{
				
				# Insert the category name from the file into the text box. 
				
						"insertText": {
							'objectId': 'SubTitleTextBox',
							'text':current_date,
							'insertionIndex':0
						}
				},
				{
				
				# Update text formatting.
				
					'updateTextStyle': {
						'objectId':'SubTitleTextBox',
						'fields': 'foregroundColor,fontFamily,fontSize',
						'style': {
							'foregroundColor': {
								'opaqueColor': {
									'rgbColor': {
										'red': 0.4,
										'green': 0.4,
										'blue': 0.4
									}
								}
							},
							'fontFamily': 'Helvetica Neue',
							'fontSize': {
								'magnitude': 30,
								'unit':'PT'
							}
						}
					}
				},
				{
				
				# Update text alignment formatting (Category Slide)
				
					'updateParagraphStyle': {
						'objectId':'SubTitleTextBox',
						'fields': 'alignment',
						'style': {
							'alignment':'END'
						}
					}
				},
				{
				
				# Create the bottom triangle.
				
					'createShape': {
						'objectId':'TitleTriangle',
						'elementProperties':
							{
								'pageObjectId': 'TitleSlide',
								'size': {
									'width': {
										'magnitude':10 * POINTS_IN_INCH,
										'unit':'PT'
									},
									'height': {
										'magnitude':2.48 * POINTS_IN_INCH,
										'unit':'PT'
									}
								},
								"transform": {
									"scaleX": 1,
									"scaleY": 1,
									"translateX": 0 * POINTS_IN_INCH,
									"translateY": 3.15 * POINTS_IN_INCH,
									"unit": "PT"
								}
							},
						'shapeType':'RIGHT_TRIANGLE',
					}
				},
				{
				
				# Formats the bottom triangle.
				
					'updateShapeProperties': {
						'objectId': 'TitleTriangle',
						'fields': 'shapeBackgroundFill,outline',
						"shapeProperties": {
							'shapeBackgroundFill': {
								'solidFill': {
									'color': {
										'rgbColor': {
											'red': 0.34,
											'green':0.73,
											'blue':0.18
										}
									}
								}
							},
							'outline': {
								'propertyState': 'NOT_RENDERED'
							}
						}
					}
				}
			]
			
	body = {
		'requests': requests
	}
	response = service.presentations().batchUpdate(presentationId=PRESENTATION_ID,body=body).execute()
	slide_position_index += 1
	
	while i < min_boundary:

		# If the category indices for a product is different, create a category slide with the corresponding category name.	
		if categories[i] != categories[i - 1]:
			requests = [
			{
				'createSlide': {
					'objectId': 'CategorySlide' + str(i),
					'insertionIndex': slide_position_index,
				}
			},
			{
				
				# Create the textbox.
					
					'createShape': {
						'objectId':'CategoryTitle' + str(i),
						'elementProperties': {
							'pageObjectId': 'CategorySlide' + str(i),
							'size': {
								'width': {
									'magnitude':9.32 * POINTS_IN_INCH,
									'unit':'PT'
								},
								'height': {
									'magnitude':1.87 * POINTS_IN_INCH,
									'unit':'PT'
								}
							},
							"transform": {
								"scaleX": 1,
								"scaleY": 1,
								"translateX": 0.34 * POINTS_IN_INCH,
								"translateY": 1.88 * POINTS_IN_INCH,
								"unit": "PT"
							}
						},
						'shapeType':'TEXT_BOX',
					}
				},
				{
				
				# Insert the category name from the file into the text box. 
				
						"insertText": {
							'objectId':'CategoryTitle' + str(i),
							'text':categories[i],
							'insertionIndex':0
						}
				},
				{
				
				# Update text formatting.
				
					'updateTextStyle': {
						'objectId':'CategoryTitle' + str(i),
						'fields': 'foregroundColor,fontFamily,fontSize',
						'style': {
							'foregroundColor': {
								'opaqueColor': {
									'rgbColor': {
										'red': 0.27,
										'green': 0.27,
										'blue': 0.27
									}
								}
							},
							'fontFamily': 'Helvetica Neue',
							'fontSize': {
								'magnitude': 48,
								'unit':'PT'
							}
						}
					}
				},
				{
				
				# Update text alignment formatting (Category Slide)
				
					'updateParagraphStyle': {
						'objectId':'CategoryTitle' + str(i),
						'fields': 'alignment',
						'style': {
							'alignment':'END'
						}
					}
				},
				{
				
				# Create the bottom rectangle.
				
					'createShape': {
						'objectId':'CategoryRectangle' + str(i),
						'elementProperties':
							{
								'pageObjectId': 'CategorySlide' + str(i),
								'size': {
									'width': {
										'magnitude':10 * POINTS_IN_INCH,
										'unit':'PT'
									},
									'height': {
										'magnitude':0.82 * POINTS_IN_INCH,
										'unit':'PT'
									}
								},
								"transform": {
									"scaleX": 1,
									"scaleY": 1,
									"translateX": 0 * POINTS_IN_INCH,
									"translateY": 4.81 * POINTS_IN_INCH,
									"unit": "PT"
								}
							},
						'shapeType':'RECTANGLE',
					}
				},
				{
				
				# Formats the bottom rectangle.
				
					'updateShapeProperties': {
						'objectId': 'CategoryRectangle' + str(i),
						'fields': 'shapeBackgroundFill,outline',
						"shapeProperties": {
							'shapeBackgroundFill': {
								'solidFill': {
									'color': {
										'rgbColor': {
											'red': 0.34,
											'green':0.73,
											'blue':0.18
										}
									}
								}
							},
							'outline': {
								'propertyState': 'NOT_RENDERED'
							}
						}
					}
				}
			]
			slide_position_index += 1
		else:
			requests = [
				{
					'createSlide': {
						'objectId': 'Slide' + str(i),
						'insertionIndex': slide_position_index,
					}
				},
				{
				
				# Create a title textbox.
					
					'createShape': {
						'objectId':'Title' + str(i),
						'elementProperties': {
							'pageObjectId': 'Slide' + str(i),
							'size': {
								'width': {
									'magnitude':9.32 * POINTS_IN_INCH,
									'unit':'PT'
								},
								'height': {
									'magnitude':0.63 * POINTS_IN_INCH,
									'unit':'PT'
								}
							},
							"transform": {
								"scaleX": 1,
								"scaleY": 1,
								"translateX": 0.34 * POINTS_IN_INCH,
								"translateY": 0.49 * POINTS_IN_INCH,
								"unit": "PT"
							}
						},
						'shapeType':'TEXT_BOX',
					}
				},
				{
				
				# Insert the titles from the file into the text box. 
				
						"insertText": {
							'objectId':'Title' + str(i),
							'text':titles[i],
							'insertionIndex':0
						}
				},
				{
				
				# Update text formatting.
				
					'updateTextStyle': {
						'objectId':'Title' + str(i),
						'fields': 'foregroundColor,fontFamily,fontSize,underline',
						'style': {
							'foregroundColor': {
								'opaqueColor': {
									'rgbColor': {
										'red': 0.27,
										'green': 0.27,
										'blue': 0.27
									}
								}
							},
							'fontFamily': 'Helvetica Neue',
							'fontSize': {
								'magnitude': 24,
								'unit':'PT'
							},
							'underline': True
						}
					}
				},	
				{
					
				# Create an image placeholder textbox.
					
					'createShape': {
						'objectId':'Image' + str(i),
						'elementProperties': {
							'pageObjectId': 'Slide' + str(i),
							'size': {
								'width': {
									'magnitude':4.37 * POINTS_IN_INCH,
									'unit':'PT'
								},
								'height': {
									'magnitude':3.74 * POINTS_IN_INCH,
									'unit':'PT'
								}
							},
							"transform": {
								"scaleX": 1,
								"scaleY": 1,
								"translateX": 0.34 * POINTS_IN_INCH,
								"translateY": 1.26 * POINTS_IN_INCH,
								"unit": "PT"
							}
						},
						'shapeType':'TEXT_BOX',
					}
				},
				{
				
				# Insert the placeholder text.
				
					"insertText": {
						'objectId':'Image' + str(i),
						'text':'placeholder',
						'insertionIndex':0
					}
				},
				{
				
				# Switch the placeholder shape with an image.
				
					"replaceAllShapesWithImage": {
						"imageUrl": images[i],
						"imageReplaceMethod": "CENTER_INSIDE",
						"pageObjectIds": [
							'Slide' + str(i)
						],
						"containsText": {
							"text":'placeholder',
							"matchCase":False
							}
						}
				},
				{
					
				# Create a product description textbox.
					
					'createShape': {
						'objectId':'ProductDesc' + str(i),
						'elementProperties': {
							'pageObjectId': 'Slide' + str(i),
							'size': {
								'width': {
									'magnitude':4.37 * POINTS_IN_INCH,
									'unit':'PT'
								},
								'height': {
									'magnitude':3.74 * POINTS_IN_INCH,
									'unit':'PT'
								}
							},
							"transform": {
								"scaleX": 1,
								"scaleY": 1,
								"translateX": 5.28 * POINTS_IN_INCH,
								"translateY": 1.26 * POINTS_IN_INCH,
								"unit": "PT"
							}
						},
						'shapeType':'TEXT_BOX',
					}
				},
				{
					"insertText": {
						'objectId':'ProductDesc' + str(i),
						'text':specs[i],
						'insertionIndex':0
					}
				},
				{
				
				# Update text formatting (Product Description).
				
					'updateTextStyle': {
						'objectId':'ProductDesc' + str(i),
						'fields': 'foregroundColor,fontFamily,fontSize',
						'style': {
							'foregroundColor': {
								'opaqueColor': {
									'themeColor':'DARK2'
								}
							},
							'fontFamily': 'Helvetica Neue',
							'fontSize': {
								'magnitude': 10,
								'unit':'PT'
							}
						}
					}
				},
				{
				
				# Update text alignment formatting (Product Description).
				
					'updateParagraphStyle': {
						'objectId':'ProductDesc' + str(i),
						'fields': 'lineSpacing',
						'style': {
							'lineSpacing': 100.0
						}
					}
				},
				{
				
				# Create a price tag.
				
					'createShape': {
						'objectId':'PriceTag' + str(i),
						'elementProperties':
							{
								'pageObjectId': 'Slide' + str(i),
								'size': {
									'width': {
										'magnitude':1.42 * POINTS_IN_INCH,
										'unit':'PT'
									},
									'height': {
										'magnitude':1.12 * POINTS_IN_INCH,
										'unit':'PT'
									}
								},
								"transform": {
									"scaleX": 1,
									"scaleY": 1,
									"translateX": 0.34 * POINTS_IN_INCH,
									"translateY": 3.78 * POINTS_IN_INCH,
									"unit": "PT"
								}
							},
						'shapeType':'ROUND_RECTANGLE',
					}
				},
				{
					"insertText": {
						'objectId':'PriceTag' + str(i),
						'text':'$' + prices[i],
						'insertionIndex':0
					}
				},
				{
				
				# Update text formatting (Price Tag).
				
					'updateTextStyle': {
						'objectId':'PriceTag' + str(i),
						'fields': 'foregroundColor,fontFamily,fontSize,italic',
						'style': {
							'foregroundColor': {
								'opaqueColor': {
									'themeColor': 'LIGHT1'
								}
							},
							'fontFamily': 'Helvetica Neue',
							'fontSize': {
								'magnitude': 16,
								'unit':'PT'
							},
							'italic': True
						}
					}
				},
				{
				
				# Update text alignment formatting (Price Tag).
				
					'updateParagraphStyle': {
						'objectId':'PriceTag' + str(i),
						'fields': 'alignment',
						'style': {
							'alignment':'CENTER'
						}
					}
				},
				{
				
				# Formats the price tag.
				
					'updateShapeProperties': {
						'objectId': 'PriceTag' + str(i),
						'fields': 'shapeBackgroundFill,outline',
						"shapeProperties": {
							'shapeBackgroundFill': {
								'solidFill': {
									'color': {
										'rgbColor': {
											'red': 1.0,
											'green':0.0,
											'blue':0.0
										}
									}
								}
							},
							'outline': {
								'outlineFill': {
									'solidFill': {
										'color': {
											'rgbColor': {
												'red': 1.0,
												'green':1.0,
												'blue':1.0
											}
										}
									}
								},
								'weight': {
									'magnitude': 4,
									'unit': 'PT'
								}
							},
							
							#TODO Create shadows underneath the price tag.
							
							'shadow': {
								'type': 'SHADOW_TYPE_UNSPECIFIED',
								'alignment': 'BOTTOM_RIGHT',
								'alpha': 0.5,
								'transform': {
									"scaleX": 100,
									"scaleY": 100,
									"shearX": 100,
									"shearY": 100,
									"translateX": 100,
									"translateY": 100,
									"unit": 'PT'
								},
								'color': {
									'rgbColor': {
										'red': 0.1,
										'green': 0.1,
										'blue': 0.1
									}
								},
								'blurRadius': {
									'magnitude': 454203985,
									'unit': 'PT'
								},
								'propertyState': 'RENDERED'
							},
							'contentAlignment': 'MIDDLE',
							
						}
					}
				},
				{
				
				# Create a seperator.
				
					'createShape': {
						'objectId':'Seperator' + str(i),
						'elementProperties':
							{
								'pageObjectId': 'Slide' + str(i),
								'size': {
									'width': {
										'magnitude':0.01 * POINTS_IN_INCH,
										'unit':'PT'
									},
									'height': {
										'magnitude':3.74 * POINTS_IN_INCH,
										'unit':'PT'
									}
								},
								"transform": {
									"scaleX": 1,
									"scaleY": 1,
									"translateX": 4.99 * POINTS_IN_INCH,
									"translateY": 1.28 * POINTS_IN_INCH,
									"unit": "PT"
								}
							},
						'shapeType':'RECTANGLE',
					}
				},
				{
				
				# Create the top rectangle.
				
					'createShape': {
						'objectId':'Rectangle' + str(i),
						'elementProperties':
							{
								'pageObjectId': 'Slide' + str(i),
								'size': {
									'width': {
										'magnitude':10 * POINTS_IN_INCH,
										'unit':'PT'
									},
									'height': {
										'magnitude':0.32 * POINTS_IN_INCH,
										'unit':'PT'
									}
								},
								"transform": {
									"scaleX": 1,
									"scaleY": 1,
									"translateX": 0 * POINTS_IN_INCH,
									"translateY": 0 * POINTS_IN_INCH,
									"unit": "PT"
								}
							},
						'shapeType':'RECTANGLE',
					}
				},
				{
				
				# Formats the top rectangle.
				
					'updateShapeProperties': {
						'objectId': 'Rectangle' + str(i),
						'fields': 'shapeBackgroundFill,outline',
						"shapeProperties": {
							'shapeBackgroundFill': {
								'solidFill': {
									'color': {
										'rgbColor': {
											'red': 0.34,
											'green':0.73,
											'blue':0.18
										}
									}
								}
							},
							'outline': {
								'propertyState': 'NOT_RENDERED'
							}
						}
					}
				},
				{
				
				# Create the top triangle.
				
					'createShape': {
						'objectId':'TopTriangle' + str(i),
						'elementProperties':
							{
								'pageObjectId': 'Slide' + str(i),
								'size': {
									'width': {
										'magnitude':10 * POINTS_IN_INCH,
										'unit':'PT'
									},
									'height': {
										'magnitude':0.32 * POINTS_IN_INCH,
										'unit':'PT'
									}
								},
								"transform": {
									"scaleX": 1,
									"scaleY": 1,
									"translateX": 0 * POINTS_IN_INCH,
									"translateY": 0.01 * POINTS_IN_INCH,
									"unit": "PT"
								}
							},
						'shapeType':'RIGHT_TRIANGLE',
					}
				},
				{
				
				# Formats the top triangle.
				# TODO Rotate the top triangle 180 degrees.
				
					'updateShapeProperties': {
						'objectId': 'TopTriangle' + str(i),
						'fields': 'shapeBackgroundFill,outline',
						"shapeProperties": {
							'shapeBackgroundFill': {
								'solidFill': {
									'color': {
										'rgbColor': {
											'red': 1.0,
											'green':1.0,
											'blue':1.0
										}
									}
								}
							},
							'outline': {
								'propertyState': 'NOT_RENDERED'
							}
						}
					}
				},
				
				# Create the bottom triangle.
				
				{
					'createShape': {
						'objectId':'BottomTriangle' + str(i),
						'elementProperties':
							{
								'pageObjectId': 'Slide' + str(i),
								'size': {
									'width': {
										'magnitude':10 * POINTS_IN_INCH,
										'unit':'PT'
									},
									'height': {
										'magnitude':0.32 * POINTS_IN_INCH,
										'unit':'PT'
									}
								},
								"transform": {
									"scaleX": 1,
									"scaleY": 1,
									"translateX": 0 * POINTS_IN_INCH,
									"translateY": 5.31 * POINTS_IN_INCH,
									"unit": "PT"
								}
							},
						'shapeType':'RIGHT_TRIANGLE',
					}
				},
				{
				
				# Formats the bottom triangle.
				
					'updateShapeProperties': {
						'objectId': 'BottomTriangle' + str(i),
						'fields': 'shapeBackgroundFill,outline',
						"shapeProperties": {
							'shapeBackgroundFill': {
								'solidFill': {
									'color': {
										'rgbColor': {
											'red': 0.34,
											'green':0.73,
											'blue':0.18
										}
									}
								}
							},
							'outline': {
								'propertyState': 'NOT_RENDERED'
							}
						}
					}
				}
			]
			slide_position_index += 1
			
		body = {
			'requests': requests
			}
				
		response = service.presentations().batchUpdate(presentationId=PRESENTATION_ID,body=body).execute()
		i += 1
		
	return i
	
def display_result_time(execution_time,num):

	# Display the total time taken to create the presentation (includes average times for each slide).
	print('\nSuccessfully created (%s) slides.' % str(num))
	print('Program took %s second(s) to execute.' % str(execution_time))
	print('Average time per slide: %s seconds(s).' % str(execution_time / num))
		
def main():

	start_time = time.time()
	titles,specs,prices,images,categories = parse_csv()
	num = create_presentation(titles,specs,prices,images,categories)
	end_time = time.time()
	display_result_time(end_time - start_time,num)
	root.mainloop()

main()