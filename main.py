import FREQ_plotter
import SCADA_reporter
import utils
import FGMO

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # print('')
    while True:
        print('Available Programs: ')
        programs = ['System Frequency Analysis', '33kV SCADA Operation Summary',
                    'FGMO Analysis', 'PDF Annotator\n']
        for i, program in enumerate(programs, 1):
            print(f'{i} : {program}')
        n = int(input('Please input the number of program you want to run:  '))

        if n == 1:
            FREQ_plotter.freq_plotter()
        elif n == 2:
            SCADA_reporter.scada_reporter()
        elif n == 3:
            FGMO.fgmo()
        elif n == 4:
            existing = input('Existing pdf file absolute path (with extension) : ')
            modified = input('Modified pdf file absolute path (with extension) : ')
            utils.pdf_decorator(existing, modified)
        msg = input('Do you want to continue [y/n]?')
        if msg.lower() == 'y':
            pass
        else:
            break

