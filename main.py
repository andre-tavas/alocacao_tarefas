"""
Created Fri Apr 7

@author: André Tavares
"""

import pandas as pd
import pulp

def load_data():
    
    xls = pd.ExcelFile('input.xlsx')
    master = pd.read_excel(xls, sheet_name = 'Input_Master')
    
    names = master['Nomes'].dropna().to_list()
    
    preferences = {}
    for name in names:
        preferences[name] = pd.read_excel(xls, sheet_name = 'Input_' + name,
                                           index_col = 0)

    xls.close()

    return master, names, preferences


def data_formatting(master, names, preferences):
    
    # Formats the tasks so there isn't blank space between words
    master['Tarefas'] = list(map(lambda x: x.title().replace(' ', ''), 
                                 master['Tarefas']))
    
    # Set of days
    days = master['Dias'].dropna().to_list()
    
    # Set of tasks
    tasks = master['Tarefas'].dropna().to_list()
    
    # Frequency of each task
    frequency = {row['Tarefas'] : row['Frequência'] 
                  for idx, row in master.iterrows()}
    
    # List to store all the possibilities of tasks throughout the days
    tasks_days = []
    
    # Dict to store the tasks weights
    p_j = {}
    
    # Stores wheter person i can execute task j
    d_ij = {}
    
    # Aux to creat the tasks_days list
    counter = dict.fromkeys(tasks, 0)
    
    for day in days:
        for task in tasks:
            
            # Creates the tasks_days list and p_j dict
            if counter[task] < frequency[task]:
                t_d = task + '_' + day
                tasks_days.append(t_d)
                p_j[t_d] = master.query('Tarefas == @task')['Peso'].values[0]
                counter[task] += 1
                
                # Creates d_ij dict
                for i in names:
                    d_ij[i, t_d] = preferences[i].loc[task, day]
    
    # Gets only the possible person-task combination
    d_ij = {key : val for key, val in d_ij.items() if val > 0}
    
    return days, tasks, tasks_days, frequency, p_j, d_ij



# Modelo
def run_model(names, tasks_days, p_j):
    '''
    - Minimize effort
    - s.t. all the tasks must be done
    
    - set of People
    - set of Tasks (ex.: louça almoço na terça)
    - p_j effort to make task j
    - d_ij preference, 0 if person i can't do task j
    
    -x_ij 1 if i is assigned to do j
    '''
    
    model = pulp.LpProblem('Alocacao_Tarefas', pulp.LpMinimize)
    
    # =========================== Variables ===================================
    
    x_ij = pulp.LpVariable.dicts('', d_ij.keys(), cat = 'Binary')
    
    z = pulp.LpVariable('z', lowBound = 0, cat = 'Continuos')
    
    mean = pulp.LpVariable('Mean', lowBound = 0, cat = 'Continuos')
    
    # ======================== Objective function =============================   
    
    model += z + pulp.lpSum(x_ij[name, task] * d_ij[name, task] for name, task in d_ij.keys())
    
    # =========================== Constraints =================================

    # (1) Each task is done only once
    for j in tasks_days:
        # People who are able to do the task j
        names = [name for name, task in d_ij.keys() if task == j]
        
        model += (
            pulp.lpSum( x_ij[name, j] for name in names) == 1
            )
        
    # (2) Defines the mean variable
    model += (
        mean == pulp.lpSum(x_ij[name, task] * p_j[task] 
                            for name, task in d_ij.keys()) / len(names)
        )
    
    # (3 and 4) Defines z
    for n in names:
        # List of tasks n is able to do
        tasks_n = [task for name, task in d_ij if name == n]
        
        model += (z >= mean - pulp.lpSum(x_ij[n, task] * p_j[task] for task in tasks_n))
        
        model += (z >= pulp.lpSum(x_ij[n, task] * p_j[task] for task in tasks_n) - mean)        
    
    # model.writeLP('Modelo.txt')
    status = model.solve(pulp.COIN_CMD(msg = True, maxSeconds=2))
    print(pulp.LpStatus[status])
    
    return model, x_ij, mean

def create_output(model, days, tasks, x_ij, d_ij, names):
    # Matrix tasks X day to store who does what
    df_final = pd.DataFrame(columns = days, index = tasks)
          
    # Calculates the weight for each person
    assigned_weight = dict.fromkeys(names, 0)
    assigned_quantity = dict.fromkeys(names, 0)
     
    for key in d_ij.keys():
        if x_ij[key].value() > 0:
            name = key[0]
            aux = key[1].split('_')
            task = aux[0]
            day = aux[1]
            
            df_final.loc[task, day] = name
            
            assigned_weight[name] += p_j[key[1]]
            assigned_quantity[name] += 1
            
    return df_final, assigned_weight, assigned_quantity
    
    

if __name__ == '__main__':
    
    master, names, preferences = load_data()
    
    global_data = data_formatting(master, 
                                  names,
                                  preferences)
    
    days, tasks, tasks_days, frequency, p_j, d_ij = global_data
    
    model, x_ij, mean = run_model(names, tasks_days, p_j)
    
    df_final, assigned_weight, assigned_quantity = create_output(model, days, tasks, x_ij, d_ij, names)
    
    # Saves final result
    where = 'alocacao_final.xlsx'
    writer = pd.ExcelWriter(where, engine='xlsxwriter')
    df_final.to_excel(writer,sheet_name='Geral')
    writer.close()
    
    
    