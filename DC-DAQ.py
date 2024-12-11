"""
DEREK CHIBUZOR
Made for AME-341b

This script is meant to be a Mac workaround for PLX-DAQ since PLX-DAQ uses COM ports (Macs do not have these).
However, this has worked on a Windows computer, so this is not exclusive to Mac.
There are two parts to this script: the port finder, and the serial reader and publisher (to Excel or CSV file).

DATA MUST BE INPUT AS
[HEADER1, DATA1, HEADER2, DATA2, HEARDER3, DATA3, ...]
"""

# If serial is not installed, type "python3 -m pip install pyserial" into a Terminal window
# Note that if you have another serial library installed, it may interfere with this one
import serial
from serial.tools import list_ports
# If xlsxwriter is not installed, type "pip3 install xlsxwriter" into a Terminal window
import xlsxwriter
# If matplotlib is not installed, type "pip3 install matplotlib" into a Terminal window
from matplotlib import pyplot as plt
# Python has a built-in datetime library
from datetime import datetime
# Python has a built-in time library
import time
# Python has a built-in traceback library
import traceback
from os import walk, path, getcwd


def main():

    # Constants
    INTERVAL_PLOT = 0.5  # Minimum number of seconds before plot is updated (semi-arbitrary)

    # This first part will find the available serial ports. The Arduino should be a USB port.
    # To be sure of the Arduino's port, run this part before and after plugging in the Arduino, and compare the
    # output. To minimize confusion, make sure no other devices are also being plugged in between the two runs.
    portList = list(list_ports.comports())
    p_ind = 0
    print("Ports:")
    for p in portList:
        print(str(p_ind) + ": " + p.device)
        p_ind += 1
    portChoice = 3# int(input("Enter the index of the port you want to use, or -1 to exit.\nChoice: "))


    if portChoice == -1:
        print("Exiting...")
    elif portChoice not in range(len(portList)):
        print("Invalid index. Exiting...")
    else:
        # This second part will actually read the serial data from the Arduino and write it to a file.
        # A live graph of the numerical data will also be generated.
        delim = ","

        # Get port info from user
        port = portList[portChoice].device
        buad = 9600#int(input("Enter the buad rate: "))


        # See the rest of serial.Serial()'s parameters here:
        # https://pyserial.readthedocs.io/en/latest/pyserial_api.html#serial.Serial.__init__
        ser = serial.Serial(port, buad)
        ser.close()
        ser.open()





        # Find the header and Arduino delay time between data
        initData = (ser.readline()).decode().rstrip('\r\n')
        initDataArr = initData.split(delim)
        print("\nMeasuring delay between Arduino data packets...")




        ser.readline()
        t1 = time.time()
        ser.readline()
        t2 = time.time()
        delayArd = t2 - t1

        graphPause = 0.5 * delayArd  # Account for data processing time
        if delayArd == 0:
            # This is unlikely to happen, but I must account for it
            print(
                "No notable Arduino delay between messages. There should be some sort of delay on the order of at least milliseconds.")
            print("If you want to see the live graph, there must be some pause for it to update.")
            print(
                "If a delay is introduced, the graph and data-writing will lag behind, but there will be no gaps in the data stream.")
            addDelay = input("Add a 1 ms delay? ('y'/'n'): ").upper() == "Y"
            if addDelay:
                delayArd = 0.001
                graphPause = 0.001
        else:
            print("Delay:", round(delayArd, 3), "s.")
        TO_time = 1.25 * delayArd  # Amount of time before timeout on serial read

        xArr = []
        yArr = []


        print(initDataArr)

        timeColIndex = 1#int(input("Enter the column index (start at 0) for the x-axis in the transmitted data: ")) #1
        dataColIndex = 3#int(input("Enter the column index (start at 0) for the y-axis in the transmitted data: ")) #3
        saveChoice = 0#int(input("Enter 0 to save as an Excel workbook, or 1 to save as a CSV file: "))#0


        ext0 = ".xlsx"
        ext1 = ".txt"

        def genfileName():
            cwd = path.abspath(getcwd())
            seed = "file"
            fileArr = []
            for (dirpath, dirnames, filenames) in walk(cwd):
                if saveChoice == 0:
                    fileArr = [x for x in filenames if ext0 and seed in x and x[len(seed)].isnumeric()]
                elif saveChoice == 1:
                    fileArr = [x for x in filenames if ext1 and seed in x and x[len(seed)].isnumeric()]
                fileArr.sort(reverse=True)
                break

            if fileArr:
                lastfileName = fileArr[0]
                lastfilext = ext0 if (ext0 in lastfileName) else ext1
                lastfileNameNumber = int(lastfileName[len(seed):len(lastfileName)-len(lastfilext)])
            else:
                lastfileNameNumber = -1

            bufferStr = "0" if (lastfileNameNumber < 9) else ""

            return seed + bufferStr + str(lastfileNameNumber+1)
        fileNameBase = genfileName()




        if saveChoice == 0:

            fileName = fileNameBase + ext0
            sheetName = "Data"
            workbook = xlsxwriter.Workbook(fileName)
            sheet = workbook.add_worksheet(sheetName)
            rowNum = 1
            format_time = workbook.add_format({'num_format': 'hh:mm:ss.000'})
            sheet.set_column(1, 1, 15)
            sheet.set_column(3, 3, 15)

            excelColNum = 0
            for colNum in range(0, len(initDataArr), 2):
                sheet.write(0, excelColNum, initDataArr[colNum])

                if initDataArr[colNum] == "TIME":
                    excelTimeColNum = excelColNum
                if initDataArr[colNum] == "DATA":
                    excelDataColNum = excelColNum

                excelColNum += 1

        elif saveChoice == 1:

            fileName = fileNameBase + ext1
            with open(fileName, "wt") as f:
                f.write("TIME,DATA,SERIAL#\n")



        fig, ax1 = plt.subplots(1, 1)
        plt.ion()

        numHeaderCols = len(initDataArr)
        ax1.set_xlabel(initDataArr[timeColIndex-1])
        ax1.set_ylabel(initDataArr[dataColIndex-1])
        currTime = None
        currData = None


        ser = serial.Serial(port, buad, timeout=delayArd)
        ser.close()
        ser.open()

        try:
            print("\nThere are three ways to stop the program:")
            print("  Press any key while the graph window is selected.")
            print("  Press the Reset button on the Arduino.")
            print("  Press Ctrl+C (use as last resort).\n")



            while True:
                dataIn = (ser.readline()).decode().rstrip('\r\n')
                dataArr = dataIn.split(delim)
                print(dataArr)

                excelColNum = 0

                for colNum in range(1, len(dataArr)+1, 2):

                    try:
                        cellData = float(dataArr[colNum])

                        if dataArr[colNum-1] == "TIME":
                            t0 = float(initDataArr[colNum])
                            cellData = (cellData - t0)/1000
                            currTime = float(cellData)

                        elif dataArr[colNum-1] == "DATA":
                            currData = cellData


                        if saveChoice == 0:
                            sheet.write(rowNum, excelColNum, cellData)

                        elif saveChoice == 1:
                            with open(fileName, "at") as f:
                                f.write(str(cellData) + ",")

                    except:
                        pass
                        currTime = None
                        currData = None
                        # Uncomment the above code you want there to be discontinuities in the graph/excel sheet when:
                        # 1) The temperature reads NAN
                        # 2) The arduino temporarily disconnects and no data is read at all for a second
                        # Else, the graph/excel sheet will plot/store the the last known value until new data is present


                    finally:
                        excelColNum += 1

                if saveChoice == 0:
                    rowNum += 1
                elif saveChoice == 1:
                    with open(fileName, "at") as f:
                        f.write("\n")


                xArr.append(currTime)
                yArr.append(currData)
                ax1.plot(xArr[-2:], yArr[-2:], '-b')
                # Singular points cannot be plotted (since it's not a scatter plot) - only an array can be plotted to generate a line
                # So, plot the newly generated point + the prior point in order to have a line.


                if plt.waitforbuttonpress(graphPause):
                    raise KeyboardInterrupt



        except KeyboardInterrupt:
            plt.savefig(fileNameBase + ".pdf", format="pdf", bbox_inches="tight")
            print("\nExiting...")
        except:
            print("\nSomething went wrong:")
            print(traceback.format_exc(), "\n")
        finally:
            ser.close()
            plt.close()


        if saveChoice == 0:
            # Create a new chart object before closing the workbook
            capitalA_Int = ord("A")
            timeCol = chr(capitalA_Int + excelTimeColNum)
            dataCol = chr(capitalA_Int + excelDataColNum)
            chartCol = chr(capitalA_Int + numHeaderCols + 1)
            chart = workbook.add_chart({'type': 'line'})
            finalRowStr = str(rowNum+1)
            chart.add_series({
                'categories': '=' + sheetName + '!$' + timeCol + '$2:$' + timeCol + '$' + finalRowStr,
                'values': '=' + sheetName + '!$' + dataCol + '$2:$' + dataCol + '$' + finalRowStr,
            })
            chart.set_x_axis({'name': '=' + sheetName + '!$' + timeCol + '$1'})
            chart.set_y_axis({'name': '=' + sheetName + '!$' + dataCol + '$1'})
            chart.set_legend({'none': True})
            # Insert the chart into the worksheet
            sheet.insert_chart(chartCol + '2', chart)
            workbook.close()

        print("Done.")


# Run main()
main()
