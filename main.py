import pandas as pd
import re
# from google.colab import files

def run_process():
  df = pd.read_excel('Conteúdos_RAW.xlsx')
  cnu = get_subjects_df(df)
  cnu.to_excel('Conteúdos.xlsx', index = False)
  # files.download('Conteúdos.xlsx')

def get_subjects_df(df: pd.DataFrame) -> pd.DataFrame:
  axis_dfs = []
  for i, not_axis in enumerate(df.iloc[:,1]):
      if i == 0:
        topics = split_topics(not_axis, True) # Conteúdos Gerais
        topics = topics[1:]
      else:
        topics = split_topics(not_axis)

      topic_dfs = []
      for topic in topics:
        subjects = split_subjects(topic)
        topic = subjects[0]
        if len(subjects) == 1:
          subjects = [subjects[0]]
        else:
          subjects = subjects[1:]
        topic_df = pd.melt(pd.DataFrame({topic: subjects}), var_name='Tópico', value_name='Conteúdo')
        topic_dfs.append(topic_df)
      axis_df = pd.concat(topic_dfs).reset_index(drop=True)
      axis_df['Eixo'] = df.iloc[i,0]
      axis_dfs.append(axis_df)

  cnu = pd.concat(axis_dfs).reset_index(drop=True)
  cnu = adjust_dataframe(cnu)
  return cnu

def adjust_dataframe(cnu: pd.DataFrame) -> pd.DataFrame:
  for character in ['.', ' ']:
    for column in ['Tópico', 'Conteúdo']:
      mask = cnu[column].str.startswith(character)
      if mask.any():
        cnu.loc[mask, column] = cnu.loc[mask, column].str[1:]

  cnu = cnu[['Eixo','Tópico','Conteúdo']]
  return cnu

def split_topics(not_axis: str, conteudos_gerais: bool = False) -> list:
  if conteudos_gerais:
    pattern = re.compile(r'\d\s[A-ZÀ-Ú]{2,}')
  else:
    pattern = re.compile(r'\.\s\d{1,2}\s[A-ZÀ-Ú][a-zà-ú ]')
  matched_indices = [match.start() for match in re.finditer(pattern, not_axis)]

  topics = []
  for i, idx in enumerate(matched_indices):
    if i == 0:
      topics.append(not_axis[:idx])
    if (i != len(matched_indices) - 1):
      topics.append(not_axis[idx:matched_indices[i+1]])
    else:
      topics.append(not_axis[idx:])

  return topics

def split_subjects(topic: str) -> list:
  pattern = re.compile(r'[ .]\d\.\d{1,2}\s')
  matched_indices = [match.start() for match in re.finditer(pattern, topic)]

  subjects = []
  if len(matched_indices) > 0:
    for i, idx in enumerate(matched_indices):
      if i == 0:
        subjects.append(topic[:idx])
      if (i != len(matched_indices) - 1):
        subjects.append(topic[idx:matched_indices[i+1]])
      else:
        subjects.append(topic[idx:])
  else:
    subjects = [topic]
  return subjects
