# 🧩 sharepoint-client-api

> A fluent TypeScript/JavaScript wrapper for interacting with SharePoint sites and admin APIs.  
> Provides a simple, type-safe interface for queries, CRUD operations, and more.

---

[![npm version](https://img.shields.io/npm/v/sharepoint-client-api?color=blue&logo=npm)](https://www.npmjs.com/package/sharepoint-client-api)
[![npm downloads](https://img.shields.io/npm/dt/sharepoint-client-api.svg?logo=npm&label=Downloads)](https://www.npmjs.com/package/sharepoint-client-api)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](./LICENSE)
![TypeScript](https://img.shields.io/badge/TypeScript-Ready-3178C6?logo=typescript)
![Bundle Size](https://img.shields.io/bundlephobia/minzip/sharepoint-client-api?label=size)

---

## ✨ Features

- 🧠 Fluent interface for building SharePoint API requests  
- 🌐 Supports **site** and **admin** endpoints  
- 🔍 Query builders: `select`, `filter`, `expand`, `orderBy`, `top`, `skip`  
- 🔧 Full CRUD support: `get`, `post`, `put`, `patch`, `delete`  
- 🚫 Optional error ignoring with `.ignore()`  
- 🔐 Easy token-based authentication  

---

## 📦 Installation

```bash
npm install sharepoint-client-api
# or
yarn add sharepoint-client-api
```

---

## 🚀 Usage

### Connect with Azure AD app credentials

```ts
import { connectWithSharePoint } from "sharepoint-client-api";

const sp = await connectWithSharePoint({
  siteHostname: "mytenant.sharepoint.com",
  tenantId: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  refreshToken: "YOUR_REFRESH_TOKEN",
  appCredentials: {
    clientId: "YOUR_CLIENT_ID",
    clientSecret: "YOUR_CLIENT_SECRET"
  }
});

// Get lists from a specific site
const lists = await sp
  .api("mysite", "lists")
  .select(["Id", "Title"])
  .filter("Title eq 'Documents'")
  .get();

console.log(lists);
```

---

### 🛠 Using admin endpoints

```ts
const adminSites = await sp
  .adminApi("/Sites('xxxx')")
  .top(5)
  .get();

console.log(adminSites);
```

---

### ⚙️ Ignoring errors

```ts
const result = await sp
  .api("mysite", "lists")
  .ignore()
  .get(); // Returns null if API fails instead of throwing
```

---

### 📨 Setting custom headers

```ts
await sp
  .api("mysite", "lists")
  .setHeaders({ "X-Custom-Header": "value" })
  .get();
```

---

## 🧱 Query Builder Methods

| Method                       | Description             |
| ---------------------------- | ----------------------- |
| `select(fields)`             | Choose specific fields  |
| `filter(condition)`          | OData filter condition  |
| `expand(fields)`             | Expand related entities |
| `orderBy(field, ascending?)` | Sort results            |
| `top(count)`                 | Limit number of results |
| `skip(count)`                | Skip results            |
| `rawQuery(params)`           | Add custom query params |

---

## 🔄 HTTP Methods

| Method        | Description            |
| ------------- | ---------------------- |
| `get()`       | Perform GET request    |
| `post(data)`  | Perform POST request   |
| `put(data)`   | Perform PUT request    |
| `patch(data)` | Perform PATCH request  |
| `delete()`    | Perform DELETE request |

---

## 📝 Notes

* Ensure `.api()` or `.adminApi()` is called before making any request.
* `.ignore()` allows safe API calls without throwing errors.
* Supports full OData query syntax for SharePoint.

---

## 📜 License

This project is licensed under the [MIT License](./LICENSE).


---
⭐ If you find this package helpful, consider giving it a star on [GitHub](https://github.com/rayen-creator/sharepoint-client-api) !
