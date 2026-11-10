---
title: "App Authentication Setup"
sidebar_position: "4"
---

# App Authentication Setup

Your Teams bot needs to authenticate with Azure to send messages. This involves configuring your Azure Bot Service and App Registration correctly.

## Authentication Methods

Choose one of the following authentication methods based on your security requirements:

1. **[Client Secret](./app-authentication/client-secret-authentication-setup.md)** - Simple password-based authentication using a client secret
2. **[User Assigned Managed Identity](./app-authentication/user-assigned-managed-identity-authentication-setup.md)** - Passwordless authentication using Azure managed identities
3. **[Federated Identity Credentials](./app-authentication/federated-identity-credentials-setup.md)** - Advanced identity federation using managed identities assigned to App Registration

Each method has different setup requirements in Azure Portal or Azure CLI.

## Sovereign Cloud

If your bot runs in a US Government (GCC-High, DoD) or China (21Vianet) cloud environment, see the Sovereign Cloud configuration guide.

## After Setup

Once you've completed the Azure setup for your chosen authentication method, you'll need to configure your application code. See the App Authentication configuration guide for details on environment variables and code configuration.

## Troubleshooting

If you encounter authentication errors, see the [Troubleshooting](troubleshooting) guide for common issues and solutions.
