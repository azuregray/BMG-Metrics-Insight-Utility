'''
------------------------------------------------------------- DOCUMENTATION STARTS HERE

[SCRIPT NAME]   UnitValuesRendering.py

[SCRIPT TYPE/DEPENDENCIES]  No External Script Dependencies; Excel Template File - "Results-ExportTemplate.xlsx"

[DESCRIPTION]   This python script combines the functionalities of both Input & Output Files Processing >
                Converting purely dimension layers of each DXF file into a single Combined Image >
                Producing rotations of the combined image for every 45 degrees interval >
                Extracting Text from each of the rotated images > Manage workspace cleanUp >
                Filter processed text > Collect all post-processed values >
                Present the results in a seamless fashion > Export to a structured Excel File on Demand.

[LIBRARIES USED]    (RUNNING HEADLESS)      ezdxf, os, cv2, shutil, easyocr, numpy, matplotlib.pyplot, PIL,
                                            difflib.SequenceMatcher, openpyxl.load_workbook, tempfile
                    (RUNNING SCRIPT AS IS)  tkinter.filedialog, time.sleep, ctypes.windll

[FUNCTIONS]     gracefulErrors(errorMessage),
                textScanner(imagePath),
                inputImageProducer(filePath),
                outputImageProducer(filePath),
                imageRotatory(imagePath, outputDir),
                inputHardCodedFilter(rawData),
                outputHardCodedFilter(rawData),
                uniquenessEngine(data),
                renderValues(dxfPath, dxfType),
                pathCorresponder(inputDirPath, outputDirPath),
                harvesterFunc(inputDirPath, outputDirPath),   --> MAIN FUNCTION
                excelProcessor(exportableData, exportDir)

[NOTE]  1.  This code runs only on DXF files.
        2.  Please make sure, filenames are exactly identical in both the folders before running this script.
        3.  Please make sure you have the Excel Template file named "Results-ExportTemplate.xlsx" in the same directory as the script.
        4.  Read the code atleast once and adjust accordingly for your machine suitability.
        5.  Make sure to have CUDA-related toolkit and libraries installed before running.
            Otherwise, EasyOCR will automatically default to CPU, which takes too much resources and time.
        6.  Main Program Code is also included for better UI/UX, independent of a Front-End Interface. (Refer section starting Line:481)

------------------------------------------------------------- DOCUMENTATION ENDS HERE
'''

import ezdxf, os, cv2, shutil, easyocr, matplotlib.pyplot as plt, tempfile
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

## Global Initialization of Progress Variables.
global progressTotal
global progressCounter

if __name__ == '__main__':
    scanner = easyocr.Reader(['en'], verbose=True)
else:
    scanner = easyocr.Reader(['en'], verbose=False)

def gracefulErrors(errorMessage, exitRequired=False):
    if exitRequired:
        if __name__ == '__main__':
            from time import sleep
            os.system('cls')
            print("\n\n:::::::: ❌ The system faced an error. ::::::::")
            sleep(0.8)
            print(f'[Error Message] ::: ⚠️ {errorMessage}')
            sleep(1)
            print("\n\n:::::::: Abort All Operations & Exit ::::::::\n")
            sleep(0.8)
            for letter in 'Press [ENTER] to exit...':
                    print(letter, flush=True, end='')
                    sleep(0.05)
            input()
            print()
            exit(1)
        else:
            raise Exception(errorMessage)
    elif not exitRequired:
        print(f'\n\n❌ WARNING ::: {errorMessage}\n\n')
        if ('errorLogs' in locals()) or ('errorLogs' in globals()):
            errorLogs.append(f'❌ WARNING ::: {errorMessage}')

def textScanner(imagePath):
    try:
        returnableData = scanner.readtext(imagePath, detail=0)
        return returnableData
    except Exception as e:
        gracefulErrors(f'Error scanning text from {imagePath}: {e}')
        return []

def inputImageProducer(filePath):
    try:
        doc = ezdxf.readfile(filePath)
        msp = doc.modelspace()
    except Exception as e:
        gracefulErrors(f'Error reading DXF file: {e}', exitRequired=True)
    
    output_folder = os.path.join(os.path.dirname(filePath), "Layers_TempDir")
    os.makedirs(output_folder, exist_ok=True)

    target_layer = "41"
    layer_entities = [entity for entity in msp if entity.dxf.layer == target_layer]
    final_output_path = os.path.join(output_folder, f'{os.path.basename(filePath)[:-4]}.png')
    
    fig, ax = plt.subplots(facecolor='black')
    ax.set_facecolor('black')
    ctx = RenderContext(doc)
    backend = MatplotlibBackend(ax)
    frontend = Frontend(ctx, backend)

    for entity in layer_entities:
        try:
            properties = ctx.resolve_all(entity)
            frontend.draw_entity(entity, properties)
        except Exception as e:
            gracefulErrors(f'Error drawing entity >> {entity.dxftype()} :: {e}')

    ax.set_aspect('equal')
    ax.autoscale()
    plt.axis('off')

    try:
        plt.savefig(final_output_path, dpi=1200, bbox_inches='tight', pad_inches=0, facecolor='black')
        plt.close(fig)
    except Exception as e:
        gracefulErrors(f'Error saving PNG for target layer >> {target_layer} :: {e}')
    
    return final_output_path

def outputImageProducer(filePath):
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = None
    
    try:
        doc = ezdxf.readfile(filePath)
        msp = doc.modelspace()
    except Exception as e:
        gracefulErrors(f'Error reading DXF file: {e}', exitRequired=True)
    
    output_folder = os.path.join(os.path.dirname(filePath), "Layers_TempDir")
    os.makedirs(output_folder, exist_ok=True)

    target_layers = {"INFTEXT61", "2", "SKVIEW2"}
    
    for layer in target_layers:
        layer_entities = [entity for entity in msp if entity.dxf.layer == layer]
        
        if not layer_entities:
            continue
        
        png_path = os.path.join(output_folder, f'{os.path.basename(filePath).replace('.dxf', '')}_{layer}.png')
        fig, ax = plt.subplots(facecolor='black')
        ax.set_facecolor('black')
        ctx = RenderContext(doc)
        backend = MatplotlibBackend(ax)
        frontend = Frontend(ctx, backend)

        for entity in layer_entities:
            try:
                properties = ctx.resolve_all(entity)
                frontend.draw_entity(entity, properties)
            except Exception as e:
                gracefulErrors(f'Error drawing entity >> {entity.dxftype()} :: {e}')

        ax.set_aspect('equal')
        ax.autoscale()
        plt.axis('off')

        try:
            plt.savefig(png_path, dpi=1200, bbox_inches='tight', pad_inches=0, facecolor='black')
            plt.close(fig)
        except Exception as e:
            gracefulErrors(f'Error saving PNG for layer >> {layer} :: {e}')
    
    combined_fig, combined_ax = plt.subplots(facecolor='black')
    combined_ax.set_facecolor('black')
    ctx = RenderContext(doc)
    backend = MatplotlibBackend(combined_ax)
    frontend = Frontend(ctx, backend)

    for entity in msp:
        if entity.dxf.layer in {"2", "SKVIEW2"}:
            try:
                properties = ctx.resolve_all(entity)
                properties.color = 7
                frontend.draw_entity(entity, properties)
            except Exception as e:
                pass

    combined_ax.set_aspect('equal')
    combined_ax.autoscale()
    plt.axis('off')
    combined_output_path = os.path.join(output_folder, "combined_layers.png")

    try:
        plt.savefig(combined_output_path, dpi=1200, bbox_inches='tight', pad_inches=0, facecolor='black')
        plt.close(combined_fig)
    except Exception as e:
        gracefulErrors(f'Error saving combined PNG :: {e}')

    all_images = []
    for file in sorted(os.listdir(output_folder)):
        if file.endswith(".png") and (file == "combined_layers.png" or "INFTEXT61" in file):
            img = Image.open(os.path.join(output_folder, file))
            all_images.append(img)

    if len(all_images) > 1:
        widths, heights = zip(*(img.size for img in all_images))
        total_width = max(widths)
        total_height = sum(heights)
        merged_image = Image.new("RGB", (total_width, total_height))
        y_offset = 0
        
        for img in all_images:
            merged_image.paste(img, (0, y_offset))
            y_offset += img.size[1]

        final_output_path = os.path.join(output_folder, "final_combined_output.png")
        merged_image.save(final_output_path)
    else:
        final_output_path = os.path.join(output_folder, "final_combined_output.png")
        if all_images:
            all_images[0].save(final_output_path)
        else:
            gracefulErrors('No layer images were found to save', exitRequired=True)

    for file in os.listdir(output_folder):
        if file != "final_combined_output.png":
            try:
                os.remove(os.path.join(output_folder, file))
            except Exception as e:
                gracefulErrors(f'Error deleting buffer image >> {file} :: {e}')
    
    return final_output_path

def imageRotatory(imagePath, outputDir):
    import numpy as np
    
    try:
        img = cv2.imread(imagePath)
        if img is None:
            gracefulErrors(f'Could not read image >> {imagePath}', exitRequired=True)

        filename_extension = os.path.splitext(os.path.basename(imagePath))
        os.makedirs(outputDir, exist_ok=True)
        
        angles = list(range(0, 316, 45)) # list(range(startAngle, EndAngle+1, AngleStepDifference))
        
        for angle in angles:
            height, width = img.shape[:2]
            center = (width / 2, height / 2)
            rotation_matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
            cos = np.abs(rotation_matrix[0, 0])
            sin = np.abs(rotation_matrix[0, 1])
            new_width = int((height * sin) + (width * cos))
            new_height = int((height * cos) + (width * sin))
            rotation_matrix[0, 2] += (new_width / 2) - center[0]
            rotation_matrix[1, 2] += (new_height / 2) - center[1]
            rotated_img = cv2.warpAffine(img, rotation_matrix, (new_width, new_height))
            output_path = os.path.join(outputDir, f'{filename_extension[0]}_rotated_{angle}{filename_extension[1]}')
            cv2.imwrite(output_path, rotated_img)

    except Exception as e:
        gracefulErrors(f'Error occurred while generating Image Rotations :: {e}', exitRequired=True)
    
    if os.path.isfile(imagePath):
        try:
            shutil.rmtree(os.path.dirname(imagePath))
        except:
            gracefulErrors(f'There was an error deleting the original unrotated image >> {imagePath}')
    
    return len(os.listdir(outputDir))

def inputHardCodedFilter(rawData):
    bufferList = [item for item in rawData if (len(item) > 2)]
    finalOutputList = []

    for item in bufferList:
        if '+' in item:
            continue
        if ('[' in item) or (']' in item):
            continue
        if '*' in item:
            continue
        item = item.replace('O', '0')
        item = item.replace('(', '')
        item = item.replace(')', '')
        item = item.replace('O', '0')
        item = item.replace(',', '.')
        item = item.replace('_', '')
        item = item.replace('-', '')
        item = item.replace("'", '')
        item = item.replace("`", '')
        item = ''.join(char for char in item if not char.islower())
        if '.' in item:
            if item.startswith('.'):
                continue
            elif item.startswith('IC'):
                try:
                    dotIndex = item.rfind('.')
                    finalOutputList.append(f'{float(item[dotIndex-2:]):.2f}')
                except:
                    finalOutputList.append(f'{float(item):.2f}')
            elif (item[0] == '0') and (item[1] != '.'):
                finalOutputList.append(f'{float(item[1:]):.2f}')
            elif item.startswith('R') and (len(item) > 1):
                finalOutputList.append(f'{float(item[1:]):.2f}')
            elif not item.endswith('.'):
                try:
                    finalOutputList.append(f'{float(item):.2f}')
                except:
                    continue

    return finalOutputList

def outputHardCodedFilter(rawData):
    bufferList = [item for item in rawData if (len(item) > 2)]
    finalOutputList = []
    
    for item in bufferList:
        item = item.replace('O', '0')
        item = item.replace('(', '')
        item = item.replace(')', '')
        item = ''.join(char for char in item if not char.islower())
        if '.' in item:
            if item.startswith('.'):
                finalOutputList.append(f'{float(item):.2f}')
            elif item.startswith('IC'):
                try:
                    dotIndex = item.rfind('.')
                    finalOutputList.append(f'{float(item[dotIndex-2:]):.2f}')
                except:
                    finalOutputList.append(f'{float(item):.2f}')
            elif (item[0] == '0') and (item[1] != '.'):
                finalOutputList.append(f'{float(item[1:]):.2f}')
            elif item.startswith('R') and (len(item) > 1):
                finalOutputList.append(f'{float(item[1:]):.2f}')
            elif not item.endswith('.'):
                try:
                    finalOutputList.append(f'{float(item):.2f}')
                except:
                    continue
    
    return finalOutputList

def uniquenessEngine(data, threshold=80):
    from difflib import SequenceMatcher
    
    result = set()
    n = len(data)
    for i in range(n):
        for j in range(i + 1, n):
            a, b = data[i], data[j]

            if isinstance(a,float) and isinstance(b, float):
                a_num = float(a)
                b_num = float(b)
                avg = (a_num + b_num) / 2
                diff_percent = abs(a_num - b_num) / avg * 100
                closeness = 100 - diff_percent
            else:
                closeness = SequenceMatcher(None, a, b).ratio() * 100

            if closeness < threshold:
                result.add(data[i])
                result.add(data[j])

    return list(result)

def renderValues(dxfPath, dxfType, progress_callback=None):
    if progress_callback:
        global progressTotal
        global progressCounter
    
    if not dxfPath:
        gracefulErrors('No DXF file selected.', exitRequired=True)
    elif not os.path.exists(dxfPath):        # DXF Input File Existence Check
        gracefulErrors(f'DXF File not found at >> {dxfPath}', exitRequired=True)
    else:
        if dxfType.lower() == 'input':
            try:
                outputImagePath = inputImageProducer(dxfPath)
            except Exception as e:
                gracefulErrors(f'There was an error at inputImageProducer() for DXF File >> {dxfPath}', exitRequired=True)
        elif dxfType.lower() == 'output':
            try:
                outputImagePath = outputImageProducer(dxfPath)
            except Exception as e:
                gracefulErrors(f'There was an error at outputImageProducer() for DXF File >> {dxfPath}', exitRequired=True)
    
    outputDir = os.path.join(os.path.dirname(dxfPath), "tempStorage")
    os.makedirs(outputDir, exist_ok=True)
    
    try:
        rotatedImagesCount = imageRotatory(outputImagePath, outputDir)
    except Exception as e:
        gracefulErrors(f'There was an error at imageRotatory() for image >> {outputImagePath}', exitRequired=True)
    
    if __name__ == '__main__':
        os.system('cls')
        if dxfType.lower() == 'input':
            print(':::::::: Processing Input File ::::::::')
        elif dxfType.lower() == 'output':
            print(':::::::: Processing Output File ::::::::')
        print(f'\n[FILE NAME] {os.path.basename(dxfPath)}')
        print(f'\n✅ Images are ready. There are in total {rotatedImagesCount} Rotated Images.\n')
    
    extracted_text_list = []
    for i, rotatedImage in enumerate(os.listdir(outputDir), start=1):
        if __name__ == '__main__':
            print(f'\nProcessing Text [Image {i}]: {rotatedImage}')
        pathToRotatedImage = os.path.join(outputDir, rotatedImage)
        
        try:
            bufferText = textScanner(pathToRotatedImage)
            extracted_text_list += bufferText
        except Exception as e:
            gracefulErrors(f'There was an issue while extracting text from image >> {pathToRotatedImage} :: {e}')
            continue
        
        if progress_callback:
            progressCounter += 1  # Increment the counter BEFORE calling the callback
            progress_callback(progressCounter/progressTotal)
        
    try:
        shutil.rmtree(outputDir)
    except:
        gracefulErrors(f'There was an issue while removing "tempStorage" working folder>> {outputDir}')
    
    if __name__ == '__main__':
        os.system('cls')
    
    sendableList = [word for sentence in extracted_text_list for word in sentence.split()]
    
    if dxfType.lower() == 'input':
        try:
            preReturnList = inputHardCodedFilter(sendableList)   # Get post filter values.
        except Exception as e:
            gracefulErrors(f'Error at inputHardCodedFilter() in renderValues() :: {e}')
            preReturnList = sendableList
    elif dxfType.lower() == 'output':
        try:
            preReturnList = outputHardCodedFilter(sendableList)   # Get post filter values.
        except Exception as e:
            gracefulErrors(f'Error at outputHardCodedFilter() in renderValues() :: {e}')
            preReturnList = sendableList
    
    returnList = uniquenessEngine(preReturnList)    # Get unique list of elements.
    
    try:
        returnList = [float(entry) for entry in returnList]
    except Exception as e:
        gracefulErrors(f'Error while converting returnList to float at renderValues() :: {e}')
    
    return returnList

def pathCorresponder(inputDirPath, outputDirPath):
    inputFilesList = sorted([file for file in os.listdir(inputDirPath) if file.lower().endswith('.dxf')])
    outputFilesList = sorted([file for file in os.listdir(outputDirPath) if file.lower().endswith('.dxf')])
    
    if (len(inputFilesList) * len(outputFilesList)) == 0:
        gracefulErrors('Trouble finding suitable files. Please make sure there are mutually named files in both directories.', exitRequired=True)
    
    mutualFilesList = [file for file in inputFilesList if file in outputFilesList]
    nonMutualFilesList = [file for file in inputFilesList if file not in outputFilesList] + [file for file in outputFilesList if file not in inputFilesList]
    
    returnableList = []
    
    for index in range(len(mutualFilesList)):
        filename = str(mutualFilesList[index])   # Use either of the list to iterate. Doesn't matter.
        inputFileName = f'{inputDirPath}/{filename}'
        outputFileName = f'{outputDirPath}/{filename}'
        flag = 'matched'
        if not (os.path.exists(inputFileName) and os.path.exists(outputFileName)):
            gracefulErrors(f'There was trouble finding corresponding file {filename} in both the selected directories.\nMake sure no program is working with the directories and wait until this program is done.')
            continue
        else:
            returnableList.append([filename, inputFileName, outputFileName, flag])
    
    for file in (nonMutualFilesList):
        returnableList.append([str(file),'unmatched'])
    
    return returnableList

def harvesterFunc(inputDirPath, outputDirPath, progress_callback=None):
    if progress_callback:
        global progressTotal
        global progressCounter
    
    finalResults = []
    corresponderList = pathCorresponder(inputDirPath, outputDirPath)
    
    if progress_callback:
        progressTotal = len(corresponderList)*16
        progressCounter = 0
    
    for fileListItem in corresponderList:
        if len(fileListItem) == 4 and fileListItem[3].lower() == 'matched':
            fileName = fileListItem[0]
            inputValuesList = renderValues(fileListItem[1], dxfType='Input', progress_callback=progress_callback)
            outputValuesList = renderValues(fileListItem[2], dxfType='Output', progress_callback=progress_callback)
            if (len(inputValuesList) * len(outputValuesList)) == 0:
                similarityVerdict = 'False'
            else:
                similarityVerdict = 'True' if all([value in inputValuesList for value in outputValuesList]) else 'False'
            inputValuesList, outputValuesList = sorted(inputValuesList, reverse=True), sorted(outputValuesList, reverse=True)
            outputValuesList = [val if val in outputValuesList else '' for val in inputValuesList] + [value for value in outputValuesList if value not in inputValuesList]
            finalResults.append([fileName[:-4], inputValuesList, outputValuesList, similarityVerdict])
        elif len(fileListItem) == 2 and fileListItem[1].lower() == 'unmatched':
            fileName = fileListItem[0]
            similarityVerdict = 'unmatched'
            finalResults.append([fileName[:-4], [], [], similarityVerdict])
            if progress_callback:
                progressCounter += 16
                progress_callback(progressCounter/progressTotal)
    
    return finalResults

def excelProcessor(exportTemplatePath, exportableData, exportDir=tempfile.gettempdir()):
    from openpyxl import load_workbook
    
    try:
        wb = load_workbook(exportTemplatePath)
    except FileNotFoundError:
        gracefulErrors(f'Excel template file not found at >> {exportTemplatePath}', exitRequired=True)
    except Exception as e:
        gracefulErrors(f'Error loading Excel template :: {e}', exitRequired=True)
    
    ws = wb["MainSheet"]
    start_row = 2

    for i, (material_id, input_vals, output_vals, verdict) in enumerate(exportableData):
        input_row = start_row + i * 2
        output_row = input_row + 1

        ws.cell(row=input_row, column=1, value=material_id)
        if (verdict.lower() == 'unmatched'):
            ws.cell(row=input_row, column=2, value=f'[{verdict.upper()}]')
        else:
            ws.cell(row=input_row, column=2, value=verdict.upper())

        if ((len(input_vals) * len(output_vals)) * (verdict.lower() != 'unmatched')):
            for j, val in enumerate(input_vals):
                ws.cell(row=input_row, column=3 + j, value=val)
            
            for j, val in enumerate(output_vals):
                if val == '':
                    continue
                ws.cell(row=output_row, column=3 + j, value=val)

    savePath = os.path.join(exportDir, "Results-EXPORT.xlsx")
    
    try:
        wb.save(savePath)
    except Exception as e:
        gracefulErrors(f'Error saving Excel file to >> {savePath} :: {e}', exitRequired=True)
    
    return savePath

if __name__ == '__main__':
    import tkinter.filedialog as fd
    from time import sleep
    from ctypes import windll
    
    windll.shcore.SetProcessDpiAwareness(1)
    errorLogs = []
    
    os.system('cls')
    print(":::::::: Let's Start ::::::::")
    sleep(1)
    
    os.system('cls')
    print(":::::::: Select a Input Files Folder to use.. in the prompt that appears now. ::::::::")
    sleep(0.8)
    inputFilesDir = fd.askdirectory(title='Select the folder with Input DXF Files.')
    os.system('cls')
    
    os.system('cls')
    print(":::::::: Select a Output Files Folder to use.. in the prompt that appears now. ::::::::")
    sleep(0.8)
    outputFilesDir = fd.askdirectory(title='Select the folder with Output DXF Files.')
    os.system('cls')
    
    print(":::::::: Scanning your folders for DXF Files... ::::::::")
    finalResults = harvesterFunc(inputFilesDir, outputFilesDir)
    
    print('\n\n:::: FINAL RESULTS ::::')
    for index, (fileName, inputValuesList, outputValuesList, similarityVerdict) in enumerate(finalResults, start=1):
            print(f'\n:::: File #{index} ::::')
            print(f'Processed File Name :::: {fileName}')
            if similarityVerdict.lower() != 'unmatched':
                print(f'Input Values :::: {inputValuesList}')
                print(f'Output Values :::: {outputValuesList}')
                if similarityVerdict == 'True':
                    print('Are they same? :::: ✅✅ YES ✅✅\n')
                elif similarityVerdict == 'False':
                    print('Are they same? :::: ❌❌ NO ❌❌\n')
            elif similarityVerdict.lower() == 'unmatched':
                print('Condition :::: ⚠️⚠️ Matching File Not Found ⚠️⚠️\n')
    
    print('\n\nData has been processed from your files.')
    
    if errorLogs:
        print('\n\n\n:::::::: But, There were also some non-critical errors. ::::::::\n')
        for index, logMessage in enumerate(errorLogs, start=1):
            print(f'{index}. {logMessage}')
    
    for letter in '\nPress [ENTER] to exit.\nOr You can also type [YES] to export the results into an Excel file.\n':
        print(letter, flush=True, end='')
        sleep(0.04)
    
    keypress = input()
    if keypress.lower() == 'yes':
        print(":::::::: Choose a folder to export the Excel in the popup that appears now. ::::::::")
        exportDir = fd.askdirectory(title='Select the folder saving your Excel Export File.')
        exportTemplatePath = os.path.dirname(__file__.replace('\\', '/')) + '/Results-ExportTemplate.xlsx'     # Script's Current Directory
        try:
            exportedPath = excelProcessor(exportTemplatePath=exportTemplatePath , exportableData=finalResults, exportDir=exportDir)
        except Exception as e:
            gracefulErrors(f'There was an issue while exporting your results at excelProcessor() :: {e}', exitRequired=True)
        print(f'\n\nYour Excel file has been exported to --> {exportedPath}.')
        print('\n\n:::::::: Now Opening your Excel Export. ::::::::')
        sleep(1)
        os.startfile(exportedPath)
    else:
        print('Thanks for using the script. Exiting Now.')
    
    print()

