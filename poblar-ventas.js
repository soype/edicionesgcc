const spreadsheet = SpreadsheetApp.openById("1IIOmkZ_rVFAtLCN6WvHU4dmoWUoVV8v7KWrTmJR0P6M");

function crearLabel(name) {
  if (typeof name !== "string" || !name.trim()) {
    Logger.log("⚠️ crearLabel called with invalid name (ignored): " + String(name));
    return null;
  }
  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}

function hasValidCurrency(body) {
  return body.includes("AR$") || body.includes("U$S");
}

function poblarVentas() {
  try {
    Logger.log("🚀 Starting poblarVentas function...");

    const sheet = spreadsheet.getSheetByName("Ventas");
    if (!sheet) {
      throw new Error('CRITICAL: Could not find "Ventas" sheet');
    }
    Logger.log('✅ "Ventas" sheet found');

    const processedLabel = crearLabel("PROCESSED_VENTAS");
    const skippedLabel = crearLabel("SKIPPED");

    // IMPORTANTE: A partir del 18/02/2026 procesamos por Label. Antes se procesaba por leido o no leido. 
    // Nunca borrar el after:2026/02/18 o se procesarán mails viejos
    const threadsSpanish = GmailApp.search('after:2026/02/18 -label:SKIPPED -label:PROCESSED_VENTAS subject:"Nuevo pedido:"');
    const threadsEnglish = GmailApp.search('after:2026/02/18 -label:SKIPPED -label:PROCESSED_VENTAS (subject:"New order:" OR subject:"You\'ve got a new order")');
    Logger.log(`📧 Found ${threadsSpanish.length} Spanish emails and ${threadsEnglish.length} English emails`);

    const threads = [...threadsSpanish, ...threadsEnglish]
      .map((thread, index) => {
        const messages = thread.getMessages();
        if (!messages || messages.length === 0) {
          throw new Error(`CRITICAL: Thread ${index} has no messages`);
        }

        const message = messages[0];
        const subject = message.getSubject();
        Logger.log(`📬 Processing email: "${subject}"`);

        const match = subject.match(/#(\d{7})/);
        const orderNumber = match ? parseInt(match[1], 10) : 0;

        if (orderNumber === 0) {
          throw new Error(`CRITICAL: No valid order number found in subject: "${subject}"`);
        }

        return { thread, orderNumber };
      })
      .filter((t) => t.orderNumber > 0);

    Logger.log(`🔢 Processing ${threads.length} valid threads with order numbers`);
    threads.sort((a, b) => a.orderNumber - b.orderNumber);

    for (let i = 0; i < threads.length; i++) {
      Logger.log(`\n🔄 Processing thread ${i + 1}/${threads.length} (Order #${threads[i].orderNumber})`);

      const messages = threads[i].thread.getMessages();
      if (!messages || messages.length === 0) {
        throw new Error("CRITICAL: No messages found in thread");
      }

      for (let j = 0; j < messages.length; j++) {
        Logger.log(`📨 Processing message ${j + 1}/${messages.length}`);

        const message = messages[j];
        if (!message) {
          throw new Error("CRITICAL: Message is null or undefined");
        }

        const body = message.getPlainBody();
        if (!body) {
          throw new Error("CRITICAL: Email body is empty or null");
        }

        if (!hasValidCurrency(body)) {
          Logger.log("⏭️ Skipping email " + body.substring(0, 200) + ": unsupported currency (no AR$ / U$S)");

          message.getThread().addLabel(skippedLabel);

          message.markRead(); 
          Logger.log("❌ Skipped email marked as read");

          continue;
        }

        Logger.log("📩 EMAIL BODY:\n" + body.substring(0, 500) + (body.length > 500 ? "...[TRUNCATED]" : ""));

        const { customerName, paymentMethod, shippingMethod, items, orderNumber, orderDate } = getOrderData(body);

        if (!items?.length) {
          throw new Error("CRITICAL: No items found to process for this order");
        }
        Logger.log(`✅ Found ${items.length} items to process`);

        if (!customerName) {
          throw new Error("CRITICAL: No customer name found in email");
        }
        Logger.log(`👤 Customer identified: ${customerName}`);

        const result = setCustomerData(body, customerName);
        const nombreCliente = result?.nombreCliente || customerName;
        Logger.log(`✅ Customer data processed: ${nombreCliente}`);

        for (let k = 0; k < items.length; k++) {
          Logger.log(`🛍️ Processing item ${k + 1}/${items.length}`);

          const item = items[k];
          if (!item) {
            throw new Error(`CRITICAL: Item ${k + 1} is null or undefined`);
          }

          const { name: itemName, quantity: itemQuantity, price: rawPrice, currency } = item;

          // Validate item data
          if (!itemName) {
            throw new Error(`CRITICAL: Item ${k + 1} has no name`);
          }
          if (!itemQuantity || itemQuantity <= 0) {
            throw new Error(`CRITICAL: Item ${k + 1} has invalid quantity: ${itemQuantity}`);
          }
          if (!rawPrice || rawPrice <= 0) {
            throw new Error(`CRITICAL: Item ${k + 1} has invalid price: ${rawPrice}`);
          }

          Logger.log(`📊 Item details: ${itemName} × ${itemQuantity} @ ${rawPrice} ${currency}`);

          const dolar = sheet.getRange("O1").getValue();
          if (!dolar || dolar <= 0) {
            throw new Error("CRITICAL: Invalid or missing dollar rate in cell O1");
          }
          Logger.log(`💱 Dollar rate: ${dolar}`);

          const priceMultiplier = currency === "U$S" ? dolar : 1;
          const itemPrice = rawPrice * priceMultiplier;
          const itemPricePerUnit = itemPrice / itemQuantity;
          const discountPercentage = 0;

          Logger.log(`💰 Calculated prices: Total=${itemPrice}, Per Unit=${itemPricePerUnit}`);

          const orderDateFormatted = orderDate ? Utilities.formatDate(orderDate, "GMT", "dd/MM/yyyy") : Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy");

          const sellDataInitial = [[orderNumber, nombreCliente, itemQuantity, itemName]];

          const sellDataFinal = [[itemPricePerUnit, discountPercentage, orderDateFormatted, shippingMethod, orderDateFormatted, orderDateFormatted, paymentMethod]];

          const columnValues = sheet.getRange("H:H").getValues();
          const lastRow = columnValues.filter(String).length;
          Logger.log(`📍 Writing to row ${lastRow + 1}`);

          sheet.getRange(lastRow + 1, 7, 1, sellDataInitial[0].length).setValues(sellDataInitial);
          sheet.getRange(lastRow + 1, 14, 1, sellDataFinal[0].length).setValues(sellDataFinal);

          Logger.log(`✅ Item #${k + 1} processed: ${itemName} × ${itemQuantity} @ ${itemPricePerUnit}`);
        }

        Logger.log("📝 Order #" + orderNumber + " processed.\n\n");

        message.markRead();
        message.getThread().addLabel(processedLabel);
        Logger.log("✅ Email marked as read");
      }
    }

    Logger.log("🏁 poblarVentas function completed");
  } catch (error) {
    Logger.log(`❌ FATAL ERROR in poblarVentas: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw error; // Re-throw to ensure hard failure
  }
}

function getOrderData(body) {
  Logger.log("🔍 Starting getOrderData analysis...");

  if (!body || typeof body !== "string") {
    throw new Error("CRITICAL: Invalid body provided to getOrderData");
  }

  const lines = body
    .split(/\r?\n/)
    .map((line) => {
      if (line === null || line === undefined) {
        Logger.log("⚠️ WARNING: Found null/undefined line in email body");
        return "";
      }
      return line.trim();
    })
    .filter(Boolean);

  Logger.log(`📄 Processed ${lines.length} non-empty lines`);

  const orderNumberMatch = lines.find((line) => line.toLowerCase().includes("order #"));
  const orderNumber = orderNumberMatch ? (orderNumberMatch.match(/#(\d+)/) || [])[1] : null;

  Logger.log(`🔢 Order number extraction: ${orderNumber ? "Found #" + orderNumber : "Not found"}`);
  if (orderNumberMatch) {
    Logger.log(`📝 Order line: "${orderNumberMatch}"`);
  }

  const customerNameMatch = body.match(/received (?:the following order|a new order) from ([\s\S]*?):/i) || body.match(/Has recibido el siguiente pedido de ([\s\S]*?):/i);

  const customerName = customerNameMatch ? customerNameMatch[1].replace(/\s+/g, " ").trim() : "";

  Logger.log(`👤 Customer name extraction: ${customerName ? '"' + customerName + '"' : "Not found"}`);

  const dateMatch = body.match(/\(\s*(\d{1,2})\s+([a-zA-Zñáéíóú]+),\s*(\d{4})\s*\)/);
  let orderDate = null;
  if (dateMatch) {
    const [_, day, monthNameRaw, year] = dateMatch;
    const monthName = monthNameRaw.toLowerCase();
    const months = {
      enero: 0,
      febrero: 1,
      marzo: 2,
      abril: 3,
      mayo: 4,
      junio: 5,
      julio: 6,
      agosto: 7,
      septiembre: 8,
      octubre: 9,
      noviembre: 10,
      diciembre: 11,
    };
    const month = months[monthName];
    if (month !== undefined) {
      orderDate = new Date(year, month, parseInt(day));
    }
  }

  const itemLines = extractItems(body);
  Logger.log(`🛍️ Processing ${itemLines.length} item lines...`);

  const items = itemLines
    .map((line, index) => {
      if (!line || typeof line !== "string") {
        throw new Error(`CRITICAL: Invalid item line ${index + 1}`);
      }

      Logger.log(`📦 Processing item line ${index + 1}: "${line}"`);

      const currency = line.includes("AR$") ? "AR$" : "U$S";
      const parts = line.split(currency);

      if (parts.length < 2) {
        throw new Error(`CRITICAL: Item line ${index + 1} missing currency separator`);
      }

      const left = parts[0] ? parts[0].trim() : "";
      const right = parts[1] ? parts[1].trim() : "";

      if (!right) {
        throw new Error(`CRITICAL: Item line ${index + 1} missing price part`);
      }

      const price = parseFloat(right.replace(".", "").replace(",", "."));

      if (isNaN(price) || price <= 0) {
        throw new Error(`CRITICAL: Item line ${index + 1} has invalid price: ${right}`);
      }

      const nameMatch = left.match(/^(.*) -/);
      const name = nameMatch ? nameMatch[1].trim() : "Desconocido";

      const quantityMatch = left.match(/[xX×]\s*(\d+)/);
      const quantity = quantityMatch ? parseInt(quantityMatch[1]) : 1;

      if (quantity <= 0) {
        throw new Error(`CRITICAL: Item line ${index + 1} has invalid quantity: ${quantity}`);
      }

      const codeMatch = left.match(/\(#([A-Z]{2,3})\s+([^)]+)\)/);
      const code = codeMatch ? `${codeMatch[1]} ${codeMatch[2]}` : "";

      Logger.log(`✅ Parsed item ${index + 1}: ${name} × ${quantity} @ ${price} ${currency}`);

      return {
        name,
        code,
        quantity,
        price,
        currency,
      };
    })
    .filter((item) => item !== null);

  Logger.log(`✅ Successfully parsed ${items.length} valid items`);

  const paymentMethod = body.includes("U$S") ? "U$S Pay Pal" : "AR$ Mercado Pago";

  let shippingMethod = "Digital";
  if (body.includes("Física Argentina") || body.includes("Castillo")) {
    shippingMethod = "Castillo";
  }

  Logger.log("🔢 Order Number: " + orderNumber);
  Logger.log("👤 Customer Name: " + customerName);
  Logger.log("📅 Order Date: " + (orderDate ? orderDate.toDateString() : "N/A"));
  Logger.log("💳 Payment Method: " + paymentMethod);
  Logger.log("🚚 Shipping Method: " + shippingMethod);
  Logger.log("🛍️ Items:");

  items.forEach((item, idx) => {
    Logger.log(`  #${idx + 1}: ${item.name} × ${item.quantity} @ ${item.price} ${item.currency} [${item.code}]`);
  });

  return {
    customerName,
    items,
    paymentMethod,
    shippingMethod,
    orderNumber,
    orderDate,
  };
}

function extractItems(body) {
  Logger.log("🔍 Starting extractItems...");

  if (!body || typeof body !== "string") {
    throw new Error("CRITICAL: Invalid body provided to extractItems");
  }

  const lines = body.split(/\r?\n/).map((line) => {
    if (line === null || line === undefined) {
      Logger.log("⚠️ WARNING: Found null/undefined line in extractItems");
      return "";
    }
    return line.trim();
  });

  const items = [];
  Logger.log(`📄 Scanning ${lines.length} lines for items...`);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line) continue;

    // Buscar línea que contiene la descripción y código del item
    if (line.match(/- Digital|digital-en|- Física Argentina/i)) {
      Logger.log(`🎯 Found item line ${i + 1}: "${line}"`);
      let itemLine = line;

      // Si la línea actual termina con "AR$" o "U$S" pero no tiene precio, buscar en la siguiente línea
      if (i + 1 < lines.length && (line.endsWith("AR$") || line.endsWith("U$S")) && !line.match(/(AR\$|U\$S)\s*[\d.,]+/)) {
        const nextLine = lines[i + 1];
        if (nextLine && nextLine.match(/^[\d.,]+$/)) {
          // La siguiente línea contiene solo números (el precio)
          itemLine = line + " " + nextLine;
          Logger.log(`🔗 Merged price from next line: "${itemLine}"`);
          i++; // avanzar para no procesar esa línea otra vez
        }
      }
      // Si la línea siguiente no tiene "AR$" o "U$S", unirla (lógica original)
      else if (i + 1 < lines.length && !line.includes("AR$") && !line.includes("U$S")) {
        const nextLine = lines[i + 1];
        if (nextLine && nextLine.match(/^(AR\$|U\$S)/)) {
          itemLine += " " + nextLine;
          Logger.log(`🔗 Merged with next line: "${itemLine}"`);
          i++; // avanzar para no procesar esa línea otra vez
        }
      }

      items.push(itemLine);
    }
  }

  Logger.log(`✅ Extracted ${items.length} item lines`);
  return items;
}

function setCustomerData(body, customerName) {
  Logger.log("👥 Starting setCustomerData...");

  if (!body || typeof body !== "string") {
    throw new Error("CRITICAL: Invalid body provided to setCustomerData");
  }

  if (!customerName) {
    throw new Error("CRITICAL: No customer name provided to setCustomerData");
  }

  const sheet = spreadsheet.getSheetByName("Datos Cliente");
  if (!sheet) {
    throw new Error('CRITICAL: Could not find "Datos Cliente" sheet');
  }

  const regex = /DNI o ID:[\s\wáéíóúñÁÉÍÓÚçã\.@´()\+°\,\-\/\*]*/;
  const match = regex.exec(body);

  if (!match) {
    Logger.log("⚠️ WARNING: No DNI/ID information found in email");
    return { nombreCliente: customerName };
  }

  Logger.log("✅ Found customer data section in email");

  const clientInfoArray = convertInfoToArray(match[0]);
  const clientInfo = clientInfoArray.filter(
    (item) =>
      item !== "" &&
      item !== "DIRECCIÓN DE FACTURACIÓN" &&
      item !== "BILLING ADDRESS" &&
      item !== "----------------------------------------" &&
      item !== "Felicitaciones por la venta." &&
      item !== "Ediciones GCC",
  );

  const cuit = clientInfo[0] && typeof clientInfo[0] === "string" ? clientInfo[0].replace("DNI o ID:", "").trim() : "";

  const nombreCliente = clientInfo[1] && typeof clientInfo[1] === "string" && clientInfo[1].trim() === customerName ? clientInfo[1].trim() : customerName;

  const coro =
    clientInfo.length === 3
      ? clientInfo[2] && typeof clientInfo[2] === "string" && clientInfo[2].trim() === customerName
        ? clientInfo[1] && typeof clientInfo[1] === "string"
          ? clientInfo[1].trim()
          : ""
        : clientInfo[2] && typeof clientInfo[2] === "string"
          ? clientInfo[2].trim()
          : ""
      : "";

  const domicilio = clientInfo
    .filter((_, index) => index > 2 && index < clientInfo.length - 2)
    .map((item) => (item && typeof item === "string" ? item.trim() : ""))
    .join("\n");

  const telCel = clientInfo[clientInfo.length - 2] && typeof clientInfo[clientInfo.length - 2] === "string" ? clientInfo[clientInfo.length - 2].trim() : "";
  const mail = clientInfo[clientInfo.length - 1] && typeof clientInfo[clientInfo.length - 1] === "string" ? clientInfo[clientInfo.length - 1].trim() : "";

  Logger.log(`📋 Customer data extracted: CUIT=${cuit}, Name=${nombreCliente}, Email=${mail}`);

  let clientExists = false;
  const dataRange = sheet.getRange("B:B");
  const data = dataRange.getValues();
  Logger.log(`🔍 Checking ${data.length} existing clients...`);

  for (let k = 1; k < data.length; k++) {
    if (data[k] && data[k][0] === nombreCliente) {
      clientExists = true;
      Logger.log(`✅ Client already exists: ${nombreCliente}`);
      break;
    }
  }

  if (!clientExists) {
    Logger.log("➕ Adding new client: " + nombreCliente);
    sheet.appendRow(["", nombreCliente, coro, mail, cuit, telCel, domicilio]);
    Logger.log(`✅ Successfully added new client: ${nombreCliente}`);
  }

  return { nombreCliente };
}

function convertInfoToArray(info) {
  if (!info || typeof info !== "string") {
    throw new Error("CRITICAL: Invalid info provided to convertInfoToArray");
  }

  return info.split("\n").map((line) => {
    if (line === null || line === undefined) {
      Logger.log("⚠️ WARNING: Found null/undefined line in convertInfoToArray");
      return "";
    }
    return line.trim();
  });
}
