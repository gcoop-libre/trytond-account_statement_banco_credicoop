# The COPYRIGHT file at the top level of this repository contains
# the full copyright notices and license terms.

from trytond.pool import Pool
from . import statement


def register():
    Pool.register(
        statement.ImportStatementStart,
        module='account_statement_banco_credicoop', type_='model')
    Pool.register(
        statement.ImportStatement,
        module='account_statement_banco_credicoop', type_='wizard')
