from os import listdir, mkdir
from os.path import split, splitext, exists, getsize
import shutil
import subprocess
import multiprocessing as mp


def full_path_in_folder(folder, file_list=None):

    if file_list is None: file_list = []

    files = listdir(folder)
    files.sort()

    for f in files:
        file_list.append(folder+f)

    return file_list


def revise_path(path):
    revised = path.replace("\\", "/")
    if revised[-1] != "/": return revised + "/"
    else: return revised

def init(EPlus_install_folder, IDFs_folder, EPWs_folder, Output_folder, Expandobject):

    global ep_folder
    global out_path
    global idfxepw
    global expobj

    ep_folder = revise_path(EPlus_install_folder)
    idf_p = revise_path(IDFs_folder)
    epw_p = revise_path(EPWs_folder)
    out_path = revise_path(Output_folder)

    if not exists(out_path):
        mkdir(out_path)

    if not exists(out_path + "temp"):
        mkdir(out_path + "temp")

    if not exists(out_path + "Warnings"):
        mkdir(out_path + "Warnings")

    if not exists(out_path + "Errors"):
        mkdir(out_path + "Errors")

    idfs = full_path_in_folder(idf_p)
    epws = full_path_in_folder(epw_p)

    idfxepw = []

    for epw in epws:
        for idf in idfs:
            idfxepw.append([idf, epw])

    expobj = Expandobject

    return idfxepw


def EPrun(idf_epw):

    idf = idf_epw[0]
    epw = idf_epw[1]

    idfname = splitext(split(idf)[1])[0]
    epwname = splitext(split(epw)[1])[0]

    local_temp_path = out_path + "temp/{}_{}/".format(idfname, epwname)
    local_outdir = out_path + "{}/".format(epwname)

    if exists(local_temp_path):
        shutil.rmtree(local_temp_path)

    mkdir(local_temp_path)

    if not exists(local_outdir):
        mkdir(local_outdir)

    shutil.copy(idf, local_temp_path + "in.idf")

    command = [ep_folder + "energyplus",  # The path of EP
               '-w', epw,  # The Path of weather file
               '-p', "eplusout",  # File name of output files
               '-d', local_temp_path,  # Output path
               '-s', 'D']

    try:

        if expobj:
            subprocess.call(ep_folder + "ExpandObjects", cwd=local_temp_path, shell=True)  # Expand objects
            if exists(local_temp_path + "expanded.idf"):
                command.append(local_temp_path + "expanded.idf") # IDF file path
            else:
                command.append(local_temp_path + "in.idf")  # IDF file path
        else:
            command.append(local_temp_path + "in.idf")  # IDF file path

        subprocess.call(command)
        subprocess.call(ep_folder + "PostProcess/ReadVarsESO", cwd=local_temp_path, shell=True) # Convert Eso into Csv

        csv_size = getsize(local_temp_path + "eplusout.csv")

        if csv_size <= 10:
            shutil.move(local_temp_path + "eplusout.err", out_path + "Errors/{}.{}.err".format(idfname, epwname))  # Move err file
            return "{} & {}: Error happened".format(idfname, epwname)

        else:
            shutil.move(local_temp_path + "eplusout.csv", local_outdir + idfname + '.csv')  # Move csv file
            shutil.move(local_temp_path + "eplusout.err", out_path + "Warnings/{}.{}.err".format(idfname, epwname))  # Move err file
            shutil.rmtree(local_temp_path)  # Delete intermediate files
            return "{} & {}: Successful".format(idfname, epwname)

    except:

        shutil.move(local_temp_path + "eplusout.err", out_path + "Errors/{}.{}.err".format(idfname, epwname))  # Move err file
        return "{} & {}: Error happened".format(idfname, epwname)


### Do not touch other codes
### Change the 4 folders here
### Besides IDF files, do not put anything else in IDFs_folder
### Besides EPW files, do not put anything else in EPWs_folder
### If the IDFs are already expanded, set Expandobject = False

queue = init(EPlus_install_folder = r"C:\EnergyPlusV8-8-0",
             IDFs_folder = r"C:\FTP\Ultimate_EP_run\IDFs",
             EPWs_folder = r"C:\FTP\Ultimate_EP_run\EPWs",
             Output_folder = r"C:\FTP\Ultimate_EP_run\Results",
             Expandobject = True)

### If you want to Adjust how many cores to use, change core_number,
### Or leave it as None, all cores will be used

core_number = None

### Do not touch other codes

if __name__ == "__main__" :

    if core_number is None:
        pool = mp.Pool()
    else:
        pool = mp.Pool(core_number)

    run_record = pool.map(EPrun, queue)

    pool.close()
    pool.join()

    print("/n")

    for rr in run_record:
        print(rr)
