# The COPYRIGHT file at the top level of this repository contains
# the full copyright notices and license terms.

from trytond.model import fields
from trytond.pool import Pool, PoolMeta
from trytond.pyson import Eval
from trytond.modules.account_statement.exceptions import ImportStatementError
from .banco_credicoop import Credicoop


class ImportStatementStart(metaclass=PoolMeta):
    __name__ = 'account.statement.import.start'

    company_party = fields.Function(
        fields.Many2One(
            'party.party', "Company Party",
            context={
                'company': Eval('company', -1),
                },
            depends={'company'}),
        'on_change_with_company_party')
    bank_account = fields.Many2One(
        'bank.account', "Bank Account",
        domain=[
            ('owners.id', '=', Eval('company_party', -1)),
            ],
        states={
            'invisible': Eval('file_format') != 'banco_credicoop',
            'required': Eval('file_format') == 'banco_credicoop',
            })

    @classmethod
    def __setup__(cls):
        super().__setup__()
        credicoop = ('banco_credicoop', 'Banco Credicoop')
        cls.file_format.selection.append(credicoop)

    @fields.depends('company')
    def on_change_with_company_party(self, name=None):
        if self.company:
            return self.company.party.id


class ImportStatement(metaclass=PoolMeta):
    __name__ = 'account.statement.import'

    def parse_banco_credicoop(self, encoding='windows-1252'):
        file_ = self.start.file_
        credicoop = Credicoop(file_)
        for bc_statement in credicoop.statements:
            statement = self.credicoop_statement(bc_statement)
            total_amount = 0
            lines_count = 0
            origins = []
            for move in bc_statement.moves:
                origins.extend(self.credicoop_origin(move))
                total_amount += move.credit - move.debit
                lines_count += 1

            statement.start_balance = 0
            statement.end_balance = total_amount
            statement.total_amount = total_amount
            statement.number_of_lines = lines_count
            statement.origins = origins
            yield statement

    def credicoop_statement(self, bc_statement):
        pool = Pool()
        Statement = pool.get('account.statement')
        Journal = pool.get('account.statement.journal')

        statement = Statement()
        statement.company = self.start.company
        journals = Journal.search([
                ('company', '=', statement.company),
                ('bank_account', '=', self.start.bank_account),
                ])
        statement.journal = journals and journals[0] or None
        if not statement.journal:
            raise ImportStatementError(
                    'To import statement, you must create a journal for '
                    'account "%(account)s".' % {
                        'account': self.start.bank_account.rec_name})

        statement.name = '%(bank_account)s@( %(date_from)s %(date_to)s )' % {
            'bank_account': self.start.bank_account.rec_name[0:24],
            'date_from': bc_statement.date_from,
            'date_to': bc_statement.date_to,
            }

        return statement

    def credicoop_origin(self, move):
        pool = Pool()
        Origin = pool.get('account.statement.origin')

        origin = Origin()
        origin.number = move.op_number
        origin.date = move.date
        origin.amount = move.credit - move.debit
        origin.description = move.concept
        origin.information = self.credicoop_information(move)
        return [origin]

    def credicoop_information(self, move):
        information = {}
        for name in ['code']:
            value = getattr(move, name)
            if value:
                information['banco_credicoop_' + name] = value
        return information
