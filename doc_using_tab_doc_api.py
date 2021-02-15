import tableaudocumentapi
from tableaudocumentapi import Datasource, xfile, Workbook, Field
from tableaudocumentapi.xfile import xml_open


# class myField (Field):
#     _ATTRIBUTES = Field._ATTRIBUTES.append('HIDDEN')


def process_datasources(datasource):
    print('data source:\t{}'.format(datasource.caption or datasource.name))
    print('version:\t{}'.format(datasource.version))
    print()
    for connection in datasource.connections:
        print('server:\t{}'.format(connection.server))
        print('dbname:\t{}'.format(connection.dbname))
        print('username:\t{}'.format(connection.username))
        print('dbclass:\t{}'.format(connection.dbclass))
        print('port:\t{}'.format(connection.port))
        print('query_band:\t{}'.format(connection.query_band))
        print('initial_sql:\t{}'.format(connection.initial_sql))
    print()
    # for field in datasource.fields.value():
    #     print(field.name)
    print('{} total fields in your data source'.format(len(datasource.fields)))

    # for count, field in enumerate(datasource.fields.values()):
    #     print ('{} {} {}'.format((count + 1), field.name, field.datatype))
    #     if field.calculation:
    #         print('      the formula is {}'.format(field.calculation))
    #     if field.default_aggregation:
    #         print('      the default aggregation is {}'.format(field.default_aggregation))
    #     if field.description:
    #         print('      the description is {}'.format(field.description))
    for field in datasource.fields.values():
        field_attributes = [field.id,
                            field.caption,
                            field.alias,
                            field.datatype,
                            field.role,
                            field.is_quantitative,
                            field.is_ordinal,
                            field.is_nominal,
                            field.calculation,
                            field.default_aggregation,
                            field.description]
        print(field_attributes)


# file_name = "C:\\Users\\jj2362\\Desktop\\docs in\\standard frequent flyer.tds"
file_name = "C:\\Users\\jj2362\\Desktop\\Sheet1 (Visual_Analytics_TOC_DataSimulated).tds"
# file_name = "C:\\Users\\jj2362\\Desktop\\docs in\\Master.twb"

file_type = xml_open(file_name)

base = file_type.getroot()
print(base.tag)

if base.tag == 'datasource':
    document = tableaudocumentapi.Datasource.from_file(file_name)
    process_datasources(document)
else:
    document = Workbook(file_name)
    for datasource in document.datasources:
        process_datasources(datasource)
        print("")

# for datasource in workbook.datasources:

test = Field()
print(test._attributes)