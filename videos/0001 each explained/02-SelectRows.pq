let
    Source = Table.FromRows({
        {"E001", "Sarah Chen", "Female", "Engineering", "2019-03-15"},
        {"E002", "Marcus Johnson", "Male", "Sales", "2020-07-22"},
        {"E003", "Emily Rodriguez", "Female", "Marketing", "2021-01-10"},
        {"E004", "David Kim", "Male", "Engineering", "2018-11-05"},
        {"E005", "Jennifer Martinez", "Female", "HR", "2019-08-20"},
        {"E006", "Robert Thompson", "Male", "Finance", "2020-02-14"},
        {"E007", "Amanda García", "Female", "Engineering", "2021-06-18"},
        {"E008", "Christopher Lee", "Male", "Operations", "2019-12-03"},
        {"E009", "Lisa Wang", "Female", "Finance", "2018-05-27"},
        {"E010", "James Wilson", "Male", "Sales", "2021-04-11"},
        {"E011", "Patricia Brown", "Female", "Marketing", "2020-09-08"},
        {"E012", "Michael Anderson", "Male", "Engineering", "2019-01-22"},
        {"E013", "Karen White", "Female", "Operations", "2021-02-16"},
        {"E014", "Daniel Martinez", "Male", "HR", "2020-11-30"},
        {"E015", "Jessica Taylor", "Female", "Sales", "2019-10-14"},
        {"E016", "Kevin Patel", "Male", "Marketing", "2021-08-25"},
        {"E017", "Michelle Clark", "Female", "Finance", "2018-09-19"},
        {"E018", "Steven Jackson", "Male", "Engineering", "2017-04-06"},
        {"E019", "Rachel Harris", "Female", "Operations", "2020-06-12"},
        {"E020", "Brian Davis", "Male", "Sales", "2019-05-28"}
    }, {"Employee ID", "Name", "Gender", "Department", "Hire date"}),   
    ChangedType = Table.TransformColumnTypes(
        Source, {
        {"Employee ID", type text},
        {"Name", type text},
        {"Gender", type text},
        {"Department", type text},
        {"Hire date", type date}
    }),
    #"Filtered Rows" = Table.SelectRows(ChangedType, each ([Gender] = "Female"))
in
    #"Filtered Rows"
