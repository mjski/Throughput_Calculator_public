############################################### -- group_quotas_sum -- #################################################
#                                                  author: Morgan D
#
# Version Python: 3.7
# Created on 28 Jan 2020
# Purpose: To take many lists of column names and a data frame, then sum if there are multiple column names in each list
#
########################################################################################################################

import re
from pandas import DataFrame


class GroupQuotasSum:

    def __init__(self):
        pass

    def quota_col_sums(self, name, df, q_list):
        if len(q_list) > 1:
            df[name] = (df.loc[:, q_list].sum(axis=1))
        elif len(q_list) == 1:
            df[name] = df[q_list]
        else:
            pass
        return df

