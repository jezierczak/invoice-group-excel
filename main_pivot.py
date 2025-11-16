from invoice_group_excel.kmexcel.pivot_tables.manager import PivotTableManager

def main() -> None:
    manager = PivotTableManager('data.xlsx')

    manager.create_pivot(
        name='FakturyAgregacja',
        source_sheet_name='Faktury',
        index=['Nrfaktury'],
        columns=['Nazwazadania'],
        values=['Kosztnieewidencjonowany','Kosztnetto'],
        agr_func='sum'

    )

    manager.create_pivot(
        name='FakturyAgregacja2',
        source_sheet_name='Faktury',
        index=['Nazwazadania'],
        # columns=['Nazwazadania'],
        values=['Kosztnieewidencjonowany','kosztbrutto','Kosztnetto'],
        agr_func='sum'

    )


    print('------------------------ [ 1 ] ------------------------')
    pivot = manager.get_pivot('FakturyAgregacja')
    print(pivot)
    manager.save_pivots_to_excel()
    #
    # print('------------------------ [ 2 ] ------------------------')
    # manager.update_pivot(
    #     name='FakturyAgregacjaSrednia',
    #     sheet_name='Faktury',
    #     index=['Nr faktury'],
    #     columns=['Nazwa zadania'],
    #     values=['Koszt netto'],
    #     aggfunc='mean'
    # )
    # pivot = manager.get_pivot('FakturyAgregacja')
    # print(pivot)
    #
    # print('------------------------ [ 3 ] ------------------------')
    # print(manager.list_pivots())
    #
    # print('------------------------ [ 4 ] ------------------------')
    # manager.remove_pivot('FakturyAgregacja')
    # try:
    #     manager.get_pivot('FakturyAgregacja')
    # except:
    #     print('Pivot not found')
    #
    # print('------------------------ [ 5 ] ------------------------')
    # manager.create_pivot(
    #     name='FakturyAgregacjaCount',
    #     sheet_name='Faktury',
    #     index=['Nr faktury'],
    #     columns=['Nazwa zadania'],
    #     values=['Koszt netto'],
    #     aggfunc='count'
    # )
    # manager.create_pivot(
    #     name='FakturyAgregacjaCountMax',
    #     sheet_name='Sprzedaz',
    #     index=['Region'],
    #     values=['Przychod'],
    #     aggfunc='max'
    # )
    #
    # manager.save_pivots_to_excel()

if __name__ == '__main__':
    main()