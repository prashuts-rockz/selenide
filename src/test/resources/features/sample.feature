Feature: Read CSV as data table

  Scenario Outline: Use csv as data table - <Scenario> , <Test> ,<Data>
    * print  'Scenario-'+'<Scenario>'

    Examples:
    |read('classpath:testData.csv')|