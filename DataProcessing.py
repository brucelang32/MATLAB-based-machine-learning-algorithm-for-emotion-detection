import pandas
import os
import numpy as np


def process_user(file_name, user_id):
    # extract user id from file name
    df = pandas.read_excel(file_name)
    # read column named 'Correct'
    correct = df['Correct']
    # calculate the average of correct
    average = correct.mean()
    res = []
    if average < 0.6:
        print("User " + user_id + " is not qualified")
        res = [user_id, average]
    else:
        HA_IN = []
        HA_OUT = []
        HA_AVPC = []
        HA_1F = []
        AN_IN = []
        AN_OUT = []
        AN_AVPC = []
        AN_1F = []
        for index, row in df.iterrows():
            if row['Response'] != 'miss' and row['Correct'] == 1 and row['randomise_blocks'] == 'AVPC' \
                    and row['Emotion'] == 'Happy' and 99 < row['Reaction Time'] < 3001:
                HA_AVPC.append(row['Reaction Time'])
            if row['Response'] != 'miss' and row['Correct'] == 1 and row['randomise_blocks'] == '1F' \
                    and row['Emotion'] == 'Happy' and 99 < row['Reaction Time'] < 3001:
                HA_1F.append(row['Reaction Time'])
            if row['Response'] != 'miss' and row['Correct'] == 1 and row['randomise_blocks'] == 'IN' \
                    and row['Emotion'] == 'Happy' and 99 < row['Reaction Time'] < 3001:
                HA_IN.append(row['Reaction Time'])
            if row['Response'] != 'miss' and row['Correct'] == 1 and row['randomise_blocks'] == 'OUT' \
                    and row['Emotion'] == 'Happy' and 99 < row['Reaction Time'] < 3001:
                HA_OUT.append(row['Reaction Time'])
            if row['Response'] != 'miss' and row['Correct'] == 1 and row['randomise_blocks'] == 'AVPC' \
                    and row['Emotion'] == 'Angry' and 99 < row['Reaction Time'] < 3001:
                AN_AVPC.append(row['Reaction Time'])
            if row['Response'] != 'miss' and row['Correct'] == 1 and row['randomise_blocks'] == '1F' \
                    and row['Emotion'] == 'Angry' and 99 < row['Reaction Time'] < 3001:
                AN_1F.append(row['Reaction Time'])
            if row['Response'] != 'miss' and row['Correct'] == 1 and row['randomise_blocks'] == 'IN' \
                    and row['Emotion'] == 'Angry' and 99 < row['Reaction Time'] < 3001:
                AN_IN.append(row['Reaction Time'])
            if row['Response'] != 'miss' and row['Correct'] == 1 and row['randomise_blocks'] == 'OUT' \
                    and row['Emotion'] == 'Angry' and 99 < row['Reaction Time'] < 3001:
                AN_OUT.append(row['Reaction Time'])

        # calculate the median of each emotion
        HA_AVPC_median = np.median(HA_AVPC)
        HA_1F_median = np.median(HA_1F)
        HA_IN_median = np.median(HA_IN)
        HA_OUT_median = np.median(HA_OUT)
        AN_AVPC_median = np.median(AN_AVPC)
        AN_1F_median = np.median(AN_1F)
        AN_IN_median = np.median(AN_IN)
        AN_OUT_median = np.median(AN_OUT)
        # calculate the superiority of each emotion
        superiority_IN = (HA_IN_median - AN_IN_median) / (HA_IN_median + AN_IN_median)
        superiority_OUT = (HA_OUT_median - AN_OUT_median) / (HA_OUT_median + AN_OUT_median)
        superiority_AVPC = (HA_AVPC_median - AN_AVPC_median) / (HA_AVPC_median + AN_AVPC_median)
        superiority_1F = (HA_1F_median - AN_1F_median) / (HA_1F_median + AN_1F_median)
        # append the result to the list
        res = [user_id, average, HA_IN_median, HA_OUT_median, HA_AVPC_median, HA_1F_median, AN_IN_median, AN_OUT_median,
               AN_AVPC_median, AN_1F_median, superiority_IN, superiority_OUT, superiority_AVPC, superiority_1F]
    return res


def data_process(file_name):
    # read .xlsx file
    df = pandas.read_excel(file_name)
    # read each line of df

    current_id = str(int(df.iloc[0]['Participant Private ID']))

    talent = []
    title = ["Experiment Version", "Participant Private ID", "Spreadsheet", "Reaction Time", "Response", "Correct",
             "Incorrect", "randomise_blocks", "randomise_trials", "Answers", "TaskImage", "Emotion"]
    talent.append(title)
    for index, row in df.iterrows():
        if str(str(row['Event Index'])) == "END OF FILE":
            # write talent to csv file
            df_talent = pandas.DataFrame(talent)
            df_talent.to_excel('./users/' + current_id + '.xlsx', index=False, header=False)
            break
        if str(int(row['Participant Private ID'])) != current_id:
            # write talent to file
            df_talent = pandas.DataFrame(talent)
            df_talent.to_excel('./users/' + current_id + '.xlsx', index=False, header=False)
            current_id = str(int(row['Participant Private ID']))
            talent = []
            talent.append(title)
        if row['display'] == 'Gender' or row['display'] == 'Age' or \
                (row['display'] == 'VisualSearchTask' and row['Response'] == "different") or (
                row['display'] == 'VisualSearchTask' and row['Response'] == "same") or (
                row['display'] == 'VisualSearchTask' and str(row['Response']) == "0"):
            res = []
            res.append(row['Experiment Version'])
            res.append(row['Participant Private ID'])
            res.append(row['Spreadsheet'])
            res.append(row['Reaction Time'])

            # res.append(row['Response'])
            response = row['Response']
            if response == 0:
                res.append('miss')
            else:
                res.append(row['Response'])

            res.append(row['Correct'])
            res.append(row['Incorrect'])

            # res.append(row['randomise_blocks'])
            blocks = row['randomise_blocks']
            if blocks == 4:
                res.append('OUT')
            elif blocks == 3:
                res.append('IN')
            elif blocks == 2:
                res.append('AVPC')
            elif blocks == 1:
                res.append('1F')

            res.append(row['randomise_trials'])
            res.append(row['Answers'])
            res.append(row['TaskImage'])

            iangeName = str(row['TaskImage'])
            if "_AN_" in iangeName:
                res.append("Angry")
            elif "_HA_" in iangeName:
                res.append("Happy")
            elif "_A_" in iangeName:
                res.append("Neutral")

            #

            talent.append(res)


if __name__ == '__main__':
    # create folder for users
    if not os.path.exists('./users'):
        os.makedirs('./users')
    if not os.path.exists('./result'):
        os.makedirs('./result')
    files = []
    for file in os.listdir('./data'):
        if file.endswith('.xlsx'):
            files.append(file)
    for file in files:
        data_process('./data/' + file)

    users_info = [['User ID', 'Correct Rate', 'HA_IN', 'HA_OUT', 'HA_AVPC', 'HA_1F', 'AN_IN', 'AN_OUT', 'AN_AVPC',
                   'AN_1F', 'Superiority_IN', 'Superiority_OUT', 'Superiority_AVPC', 'Superiority_1F']]
    files = []
    for file in os.listdir('./users'):
        if file.endswith('.xlsx'):
            files.append(file)
    for file in files:
        user_id = file.split('.')[0]
        filename = './users/' + file
        user_info = process_user(filename, user_id)
        users_info.append(user_info)
    df_talent = pandas.DataFrame(users_info)
    df_talent.to_excel('./result/users_info.xlsx', index=False, header=False)
