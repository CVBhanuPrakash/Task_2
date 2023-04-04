import openpyxl as openpyxl
import pandas as pd
from flask import Flask, request, render_template, send_file

school_list = []
class_number = []
section = []
test_no = []
output = []
for_csv = []
df = []
df_2 = []
app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])
def upload():
    global  df,df_2,school_list,class_number, test_no, for_csv
    if request.method == 'POST':
        school_list.clear()
        class_number.clear()
        section.clear()
        test_no.clear()
        df = pd.read_excel(request.files.get('file'),header=1, skipfooter=0)
        # df = df.set_index(['Name', 'Email', 'Class', 'Section', 'School Name'])

        df.drop(['SerialNo', 'Section','School Name'], axis = 1, inplace = True)
        df.Email = df.Email.str.strip("@username.com")

        df.columns = df.columns.str.replace(' - ', '$')
        df.columns = df.columns.str.replace('- ', '$')
        df.columns = df.columns.str.replace(' -', '$')

        head_split = df.columns.str.split("$", expand=True).values

        df.columns = pd.MultiIndex.from_tuples([('Basic_Info', x[0]) if pd.isnull(x[1]) else x for x in head_split])

        df = df.set_index("Basic_Info")
        df_1 = df.stack(0).reset_index(1)
        df_1.columns = df_1.columns.str.replace('level_1', 'Test_Name')
        df_2 = df_1.reset_index()


        Name, Username,Class = zip(*df_2.Basic_Info)
        df_2.Basic_Info = Name
        df_2.columns = df_2.columns.str.replace('Basic_Info', 'Name')
        df_2.insert(loc = 1, column = 'Username', value = Username )
        df_2.insert(loc = 2, column = 'Class', value = Class )

        df_2 = df_2[df_2.score != '-']
        out_file_name = "db_output_" + 'Test Data - Task 2.xlsx' + ".csv"

        df_2.to_csv(out_file_name, index=False)
        for i in range(0,len(df_2)):
             if df_2.iloc[i,2] not in class_number:
                class_number.append(df_2.iloc[i,2])
             if df_2.iloc[i,3] not in test_no:
                test_no.append(df_2.iloc[i,3])
        class_number.sort()
        test_no.sort()
        print(class_number)
        return render_template('index.html',classlists = class_number, test_numb = test_no)
    return render_template('index.html', classlists = class_number, test_numb = test_no)

@app.route("/table" , methods=['GET', 'POST'])
def table():
    for_csv.clear()
    global df_final
    class_number_selected = request.form.get('classes')
    test_selected = request.form.get('test_number')
    output = []
    for i in range(0,len(df_2)):
        if class_number_selected == str(df_2.iloc[i,2]):
            if test_selected == str(df_2.iloc[i,3]):
                output.append([df_2.iloc[i,0],df_2.iloc[i,1],df_2.iloc[i,2],df_2.iloc[i,6],df_2.iloc[i,4],df_2.iloc[i,5],df_2.iloc[i,9],df_2.iloc[i,7],df_2.iloc[i,8]])
                for_csv.append([df_2.iloc[i,0],df_2.iloc[i,1],df_2.iloc[i,2],df_2.iloc[i,6],df_2.iloc[i,4],df_2.iloc[i,5],df_2.iloc[i,9],df_2.iloc[i,7],df_2.iloc[i,8]])
    df_final = pd.DataFrame(for_csv, columns = ['Student Name', 'Student Id', 'Class','Score','Answered','Correct','Wrong','Skipped','Time taken'])
    return render_template('table.html', stu_details = output,class_number = class_number_selected,test_select = test_selected )

@app.route('/download')
def download_file():
    global for_csv
    print(for_csv)
    df_final = pd.DataFrame(for_csv, columns = ['Student Name', 'Student Id', 'Class','Score','Answered','Correct','Wrong','Skipped','Time taken'])
    df_final.to_csv('finaloutput.csv')
    path = "finaloutput.csv"
	#path = "sample.txt"
    return send_file(path, as_attachment=True)
    
if __name__ == '__main__':
    app.run(debug=True)