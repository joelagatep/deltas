import pandas as pd
import warnings

def main():
    
    pd.set_option('display.max_columns', None)
    pd.set_option("display.max_rows", None, "display.max_columns", 6)

    output_xls = r"whats_new_deltas.xlsx"

    fn_after = "Whats_New_in_Workday_2022R1_20220121_changed"
    fn_bfore = "Whats_New_in_Workday_2022R1_20220121"

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        df_bfore = pd.read_excel(fn_bfore + ".xlsx", "Whats New in Workday", header=0, engine="openpyxl")

    df_bfore.columns = [c.replace(" #", "") for c in df_bfore.columns]
    df_bfore.columns = [c.replace("'", "") for c in df_bfore.columns]
    df_bfore.columns = [c.replace("(", "") for c in df_bfore.columns]
    df_bfore.columns = [c.replace(")", "") for c in df_bfore.columns]
    
    df_bfore = df_bfore[["Whats New Item", \
                         "Feature", \
                         "Feature Description", \
                         "Community Post", \
                         "Functional Areas", \
                         "Setup Effort", \
                         "New Functionality Title", \
                         "New Functionality", \
                         "Training & Testing Impact", \
                         "Tenant", \
                         "Tentative Production Date", \
                         "Business Processes", \
                         "Tasks", \
                         "Web Services", \
                         "Security Domains", \
                         "Current Name", \
                         "Former Name", \
                         "JIRA" ]]
    df_bfore.fillna('', inplace=True)    
    
    df_bfore = df_bfore.sort_values(by=["JIRA", "New Functionality Title"])

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        df_after = pd.read_excel(fn_after + ".xlsx", "Whats New in Workday", header=0, engine="openpyxl")

    df_after.columns = [c.replace(" #", "") for c in df_after.columns]
    df_after.columns = [c.replace("'", "") for c in df_after.columns]
    df_after.columns = [c.replace("(", "") for c in df_after.columns]
    df_after.columns = [c.replace(")", "") for c in df_after.columns]
    
    df_after = df_after[["Whats New Item", \
                         "Feature", \
                         "Feature Description", \
                         "Community Post", \
                         "Functional Areas", \
                         "Setup Effort", \
                         "New Functionality Title", \
                         "New Functionality", \
                         "Training & Testing Impact", \
                         "Tenant", \
                         "Tentative Production Date", \
                         "Business Processes", \
                         "Tasks", \
                         "Web Services", \
                         "Security Domains", \
                         "Current Name", \
                         "Former Name", \
                         "JIRA" ]]
    df_after.fillna('', inplace=True)
    
    df_after = df_after.sort_values(by=["JIRA", "New Functionality Title"])

    writer = pd.ExcelWriter(output_xls, engine = 'xlsxwriter')
    
    df_new = pd.DataFrame()
    df_chg = pd.DataFrame()
    
    list_new = []
    list_chg = []
    list_headers = ["Whats New Item", \
                    "Feature", \
                    "Feature Description", \
                    "Community Post", \
                    "Functional Areas", \
                    "Setup Effort", \
                    "New Functionality Title", \
                    "New Functionality", \
                    "Training & Testing Impact", \
                    "Tenant", \
                    "Tentative Production Date", \
                    "Business Processes", \
                    "Tasks", \
                    "Web Services", \
                    "Security Domains", \
                    "Current Name", \
                    "Former Name", \
                    "JIRA"]

    for Index_Of_After in df_after.index:
    
        #
        # Find rows in After that is not in Bfore
        #
        
        df_existing = df_bfore[(df_bfore['JIRA']==df_after['JIRA'][Index_Of_After]) \
                          & (df_bfore['New Functionality Title']==df_after['New Functionality Title'][Index_Of_After])]
        list_existing = (df_bfore[(df_bfore['JIRA']==df_after['JIRA'][Index_Of_After]) \
                          & (df_bfore['New Functionality Title']==df_after['New Functionality Title'][Index_Of_After])]).values.tolist()
        
        if df_existing.empty:

            #
            # Place in New sheet.
            # This is a new row. Place in original order.
            #
            list_new.append([df_after['Whats New Item'][Index_Of_After], \
                             df_after['Feature'][Index_Of_After], \
                             df_after['Feature Description'][Index_Of_After], \
                             df_after['Community Post'][Index_Of_After], \
                             df_after['Functional Areas'][Index_Of_After], \
                             df_after['Setup Effort'][Index_Of_After], \
                             df_after['New Functionality Title'][Index_Of_After], \
                             df_after['New Functionality'][Index_Of_After], \
                             df_after['Training & Testing Impact'][Index_Of_After], \
                             df_after['Tenant'][Index_Of_After], \
                             df_after['Tentative Production Date'][Index_Of_After], \
                             df_after['Business Processes'][Index_Of_After], \
                             df_after['Tasks'][Index_Of_After], \
                             df_after['Web Services'][Index_Of_After], \
                             df_after['Security Domains'][Index_Of_After], \
                             df_after['Current Name'][Index_Of_After], \
                             df_after['Former Name'][Index_Of_After], \
                             df_after['JIRA'][Index_Of_After] \
                             ])
        
        else:  # Check if there is a change in any of the fields except for the indexes

            Changes = ""
            for Index_Of_Header, Field in enumerate(list_headers):
                if list_existing[0][Index_Of_Header] != df_after[Field][Index_Of_After]:
                    if Changes == "":
                        Changes = Changes + Field
                    else:
                        Changes = Changes + ", " + Field
                    # endif
                # endif
            # endfor
            
            if Changes != "":  # There have been some changes for this existing row
                list_chg.append([Changes, \
                                 df_after['Whats New Item'][Index_Of_After], \
                                 df_after['Feature'][Index_Of_After], \
                                 df_after['Feature Description'][Index_Of_After], \
                                 df_after['Community Post'][Index_Of_After], \
                                 df_after['Functional Areas'][Index_Of_After], \
                                 df_after['Setup Effort'][Index_Of_After], \
                                 df_after['New Functionality Title'][Index_Of_After], \
                                 df_after['New Functionality'][Index_Of_After], \
                                 df_after['Training & Testing Impact'][Index_Of_After], \
                                 df_after['Tenant'][Index_Of_After], \
                                 df_after['Tentative Production Date'][Index_Of_After], \
                                 df_after['Business Processes'][Index_Of_After], \
                                 df_after['Tasks'][Index_Of_After], \
                                 df_after['Web Services'][Index_Of_After], \
                                 df_after['Security Domains'][Index_Of_After], \
                                 df_after['Current Name'][Index_Of_After], \
                                 df_after['Former Name'][Index_Of_After], \
                                 df_after['JIRA'][Index_Of_After] \
                                 ])
            # endif
        # endif - check if Index_Of_After exists
        
    # endfor

    df_new = pd.DataFrame(list_new, columns = [ \
                             'Whats New Item', \
                             'Feature', \
                             'Feature Description', \
                             'Community Post', \
                             'Functional Areas', \
                             'Setup Effort', \
                             'New Functionality Title', \
                             'New Functionality', \
                             'Training & Testing Impact', \
                             'Tenant', \
                             'Tentative Production Date', \
                             'Business Processes', \
                             'Tasks', \
                             'Web Services', \
                             'Security Domains', \
                             'Current Name', \
                             'Former Name', \
                             'JIRA' \
                             ])

    df_chg = pd.DataFrame(list_chg, columns = [ \
                             'What Changed?', \
                             'Whats New Item', \
                             'Feature', \
                             'Feature Description', \
                             'Community Post', \
                             'Functional Areas', \
                             'Setup Effort', \
                             'New Functionality Title', \
                             'New Functionality', \
                             'Training & Testing Impact', \
                             'Tenant', \
                             'Tentative Production Date', \
                             'Business Processes', \
                             'Tasks', \
                             'Web Services', \
                             'Security Domains', \
                             'Current Name', \
                             'Former Name', \
                             'JIRA' \
                             ])

    tab = 'New Additions'
    df_new.to_excel(writer, sheet_name = tab, index=False, header=False, startrow=1)
    for column in df_new:
        column_width = max(df_new[column].astype(str).map(len).max(), len(column))
        col_idx = df_new.columns.get_loc(column)
        writer.sheets[tab].set_column(col_idx, col_idx, column_width)    
    column_settings1 = [{'header': column} for column in df_new.columns]
    (max_row, max_col) = df_new.shape
    worksheet1 = writer.sheets[tab]
    worksheet1.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings1})
    worksheet1.freeze_panes(1, 0)

    tab = 'New Changes to Previous'
    df_chg.to_excel(writer, sheet_name = tab, index=False, header=False, startrow=1)
    for column in df_chg:
        column_width = max(df_chg[column].astype(str).map(len).max(), len(column))
        col_idx = df_chg.columns.get_loc(column)
        writer.sheets[tab].set_column(col_idx, col_idx, column_width)    
    column_settings1 = [{'header': column} for column in df_chg.columns]
    (max_row, max_col) = df_chg.shape
    worksheet1 = writer.sheets[tab]
    worksheet1.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings1})
    worksheet1.freeze_panes(1, 1)

    writer.save()

if __name__ == '__main__':
    main()
