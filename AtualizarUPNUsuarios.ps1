<#
    Criado por: Rafael Carvalho
    GitHub: https://github.com/RaffaelCarv/PowerShell
    Criacao: 21 de julho de 2022
    Ultima atualizacao: 21 de julho de 2025

    Descricao:
    Este script realiza duas operacoes principais sobre contas de usuario no Active Directory:

    1. Atualiza os atributos de pais (c, co, countryCode) para padrao brasileiro.
    2. Altera o sufixo UPN dos usuarios (a parte apos o @ no UserPrincipalName), com base em um sufixo atual informado e um novo sufixo desejado.

    Funcionamento:
    - O script solicita ao administrador que informe o sufixo UPN atual (ex: @contoso.local) e o novo sufixo desejado (ex: @contoso.com).
    - Todos os usuarios com o sufixo atual informado terao o UPN ajustado para usar o novo sufixo.
    - O script tambem pergunta se deseja apenas simular as alteracoes (modo WhatIf).
    - Um log detalhado e gerado no Desktop somente se as alteracoes forem reais.

    Termos utilizados:
    - Sufixo UPN: parte do UserPrincipalName que vem apos o caractere @ (ex: em usuario@contoso.local, o sufixo UPN e "@contoso.local").
    - SearchBase: caminho LDAP que define o escopo da busca no AD (ex: OU=Usuarios,DC=empresa,DC=com).

    Exemplos:
    - Sufixo UPN atual: @contoso.local
    - Novo sufixo desejado: @contoso.com
#>
