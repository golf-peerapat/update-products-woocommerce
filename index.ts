import { Elysia, t } from "elysia";
import { read, utils, WorkBook, WorkSheet } from "xlsx";
import dotenv from "dotenv";

type RowData = {
  ID: string;
  type: string;
  SKU: string;
  Name: string;
  shortDescription?: string;
  description?: string;
  stock?: number;
  salePrice?: number;
  regularPrice?: number;
  weight?: number;
  length?: number;
  width?: number;
  height?: number;
  categories?: string;
  image?: string;
  parent?: string;
  swatchesAttributes?: string;
  brand?: string;
  attribute?: string;
  attributeValue?: string;
  installment_variable?: string;
  rtwpvg_images?: string;
  json?: string;
  "Meta: is_installment_variable_attributes"?: string;
  "Meta: rtwpvg_images"?: string;
  lazada_product_id?: string;
};
type AttributeData = {
  name: string;
  color: string;
  image: string;
  show_tooltip: string;
  tooltip_text: string;
  tooltip_image: string;
  image_size: string;
};

let skuimgResult: RowData[] = [];
let attributeResult: RowData[] = [];
let jsonResult: RowData[] = [];
type JsonType = Record<string, AttributeData>;
let jsonType: JsonType[] = [];
dotenv.config();

new Elysia().onRequest(({ set, request }) => {
  const allowedOrigins = [
    "https://product.65smarttools.com", "http://localhost:5173"
  ];
  const origin = request.headers.get("origin") || "";
  if (allowedOrigins.includes(origin)) {
    set.headers["Access-Control-Allow-Origin"] = origin;
  }
  set.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS";
  set.headers["Access-Control-Allow-Headers"] = "Content-Type";
}).post(
  "/skuimg",
  async ({ body: { upload } }) => {
    if (!upload) {
      throw new Error("No file uploaded");
    }
    if (!upload.name.endsWith(".xlsx")) {
      throw new Error("Invalid file type. Please upload an XLSX file.");
    }
    if (upload.size > 25 * 1024 * 1024) {
      throw new Error("File is too large. Maximum size is 25 MB.");
    }
    if (!upload.name.startsWith("skuimg")) {
      throw new Error("ไฟล์ไม่ถูกต้อง จะต้องเป็นไฟล์ 'skuimg......xlsx'");
    }
    try {
      const data = await upload.arrayBuffer();
      const wb: WorkBook = read(data);
      const sheet: WorkSheet = wb.Sheets[wb.SheetNames[0]];
      const jsonData: (string | null)[][] = utils.sheet_to_json(sheet, {
        header: 1,
        range: 4,
      });
      const seen = new Map<string, RowData>();
      const variations: RowData[] = [];
      const productMap = new Map<string, (string | null)[][]>();
      jsonData.forEach((row) => {
        const productId = row[0] as string | null;
        if (productId) {
          if (!productMap.has(productId)) {
            productMap.set(productId, []);
          }
          productMap.get(productId)!.push(row);
        }
      });
      productMap.forEach((rows, productId) => {
        const firstRow = rows[0];
        const nameProduct = firstRow[2] as string | null;
        const imageData = firstRow.slice(7, 15).filter(Boolean).join(",");
        const isVariable = rows.length > 1;
        if (isVariable) {
          const allAttributes = rows.map((r) => r[16] || "ไม่มีตัวเลือก");
          const combinedAttributes = [...new Set(allAttributes)].join(",");
          seen.set(productId, {
            ID: "",
            type: "variable",
            SKU: "",
            Name: nameProduct ?? "",
            image: imageData,
            attributeValue: combinedAttributes,
            lazada_product_id: productId,
          });
          rows.forEach((r) => {
            const skuProduct = r[15] as string | null;
            const variationsCombo = r[16] || "ไม่มีตัวเลือก";
            variations.push({
              ID: "",
              type: "variation",
              SKU: skuProduct ?? "",
              Name: `${nameProduct ?? ""} - ${variationsCombo}`,
              image: r[7] ?? "",
              attributeValue: variationsCombo,
              installment_variable: "yes",
              rtwpvg_images: r.slice(7, 15).filter(Boolean).join(","),
              lazada_product_id: productId ?? "",
            });
          });
        } else {
          const skuProduct = firstRow[15] as string | null;
          seen.set(productId, {
            ID: "",
            type: "simple",
            SKU: skuProduct ?? "",
            Name: nameProduct ?? "",
            image: imageData,
            attributeValue: firstRow[16] ?? "",
            lazada_product_id: productId,
          });
        }
      });
      skuimgResult = [...Array.from(seen.values()), ...variations];
      const newSheet = utils.json_to_sheet(skuimgResult);
      const newWb = utils.book_new();
      utils.book_append_sheet(newWb, newSheet, "Result");
      const csv = utils.sheet_to_csv(newSheet);
      return csv;
    } catch (error) {
      console.error("❌ Error processing skuimg file:", error);
      throw new Error("Failed to process skuimg file. Please try again.");
    }
  },
  {
    body: t.Object({ upload: t.File() }),
  }
).post(
  "/basic",
  async ({ body: { upload } }) => {
    if (skuimgResult.length === 0) {
      throw new Error("No data from skuimg to process.");
    }
    if (!upload) throw new Error("No file uploaded");
    const data = await upload.arrayBuffer();
    const wb: WorkBook = read(data);
    const sheet: WorkSheet = wb.Sheets[wb.SheetNames[0]];
    const basicData: (string | null)[][] = utils.sheet_to_json(sheet, {
      header: 1,
      range: 4,
    });
    const descriptionMap = new Map<string, string>();
    const shortDescriptionMap = new Map<string, string>();
    const imageMap = new Map<string, string>();
    basicData.forEach((row) => {
      const productId = row[0] as string | null;
      const shortDescription = row[20] as string | null;
      const description = row[19] as string | null;
      const imageList = row.slice(5, 13).filter(Boolean).join(",");
      if (productId && description) {
        descriptionMap.set(productId, description);
      }
      if (productId && shortDescription) {
        shortDescriptionMap.set(productId, shortDescription);
      }
      if (productId && imageList) {
        imageMap.set(productId, imageList);
      }
    });
    skuimgResult = skuimgResult.map((item) => {
      const isVariation = item.type === "variation";
      return {
        ...item,
        shortDescription: isVariation ? "" : shortDescriptionMap.get(item.lazada_product_id as string) || "",
        description: isVariation ? "" : descriptionMap.get(item.lazada_product_id as string) || "",
        image: isVariation ? item.image : imageMap.get(item.lazada_product_id as string) || item.image,
      };
    });
    const newSheet = utils.json_to_sheet(skuimgResult);
    const newWb = utils.book_new();
    utils.book_append_sheet(newWb, newSheet, "Result");
    const csv = utils.sheet_to_csv(newSheet);
    return csv;
  },
  {
    body: t.Object({ upload: t.File() }),
  }
).post(
  "/attribute",
  async ({ body: { upload } }) => {
    if (skuimgResult.length === 0) throw new Error("No data to process.");
    if (!upload) throw new Error("No file uploaded");
    const data = await upload.arrayBuffer();
    const wb: WorkBook = read(data);
    const tabNameMap = new Map<string, string>();
    const brandMap = new Map<string, string>();
    wb.SheetNames.forEach((sheetName) => {
      const sheet: WorkSheet = wb.Sheets[sheetName];
      const data: (string | null)[][] = utils.sheet_to_json(sheet, {
        header: 1,
        range: 4,
      });
      data.forEach((row) => {
        const productId = row[0] as string | null;
        const brand = row[3] as string | null;
        if (productId) {
          tabNameMap.set(productId, sheetName);
          if (brand) {
            brandMap.set(productId, brand);
          }
        }
      });
    });
    skuimgResult = skuimgResult.map((item) => {
      return {
        ...item,
        categories: tabNameMap.get(item.lazada_product_id as string) || "",
        brand: brandMap.get(item.lazada_product_id as string) || "",
      };
    });
    const newSheet = utils.json_to_sheet(skuimgResult);
    const newWb = utils.book_new();
    utils.book_append_sheet(newWb, newSheet, "Result");
    const csv = utils.sheet_to_csv(newSheet);
    return csv;
  },
  {
    body: t.Object({ upload: t.File() }),
  }
).post(
  "/pricestock",
  async ({ body: { upload } }) => {
    if (skuimgResult.length === 0) throw new Error("No data to process.");
    if (!upload) throw new Error("No file uploaded");
    const data = await upload.arrayBuffer();
    const wb: WorkBook = read(data);
    const sheet: WorkSheet = wb.Sheets[wb.SheetNames[0]];
    const priceStockData: (string | null)[][] = utils.sheet_to_json(sheet, {
      header: 1,
      range: 4,
    });
    const stockMap = new Map<string, number>();
    const priceMap = new Map<string, number>();
    const salePriceMap = new Map<string, number>();
    priceStockData.forEach((row) => {
      const sku = row[11] as string | null;
      const stock = [row[12], row[13], row[14], row[15], row[16]]
        .map((value) => {
          const num = Number(value);
          return isNaN(num) ? 0 : num;
        })
        .reduce((sum, value) => sum + value, 0);

      const price = Number(row[10]);
      const salePrice = Number(row[7]);
      if (sku) {
        stockMap.set(sku, stock);
        priceMap.set(sku, price);
        salePriceMap.set(sku, salePrice);
      }
    });
    skuimgResult = skuimgResult.map((item) => ({
      ...item,
      stock: stockMap.get(item.SKU),
      salePrice: salePriceMap.get(item.SKU),
      regularPrice: priceMap.get(item.SKU),
    }));
    const newSheet = utils.json_to_sheet(skuimgResult);
    const newWb = utils.book_new();
    utils.book_append_sheet(newWb, newSheet, "Result");
    const csv = utils.sheet_to_csv(newSheet);
    return csv;
  },
  {
    body: t.Object({ upload: t.File() }),
  }
).post(
  "/freight",
  async ({ body: { upload } }) => {
    if (skuimgResult.length === 0) throw new Error("No data to process.");
    if (!upload) throw new Error("No file uploaded");
    const data = await upload.arrayBuffer();
    const wb: WorkBook = read(data);
    const sheet: WorkSheet = wb.Sheets[wb.SheetNames[0]];
    const freightData: (string | null)[][] = utils.sheet_to_json(sheet, {
      header: 1,
      range: 4,
    });
    const weightMap = new Map<string, number>();
    const lengthMap = new Map<string, number>();
    const widthMap = new Map<string, number>();
    const heightMap = new Map<string, number>();
    freightData.forEach((row) => {
      const sku = row[7] as string | null;
      const weight = Number(row[6]);
      const length = Number(row[8]);
      const width = Number(row[9]);
      const height = Number(row[10]);
      if (sku) {
        weightMap.set(sku, weight);
        lengthMap.set(sku, length);
        widthMap.set(sku, width);
        heightMap.set(sku, height);
      }
    });
    skuimgResult = skuimgResult.map((item) => ({
      ...item,
      weight: weightMap.get(item.SKU),
      length: lengthMap.get(item.SKU),
      width: widthMap.get(item.SKU),
      height: heightMap.get(item.SKU),
    }));
    const newSheet = utils.json_to_sheet(skuimgResult);
    const newWb = utils.book_new();
    utils.book_append_sheet(newWb, newSheet, "Result");
    const csv = utils.sheet_to_csv(newSheet);
    return csv;
  },
  {
    body: t.Object({ upload: t.File() }),
  }
).post(
  "/wc-product-export",
  async ({ body: { upload } }) => {
    function fillDefaultFields(base: Partial<RowData>, fallback: Partial<RowData> = {}): RowData {
      return {
        ...base,
        shortDescription: base.shortDescription ?? fallback.shortDescription ?? "",
        description: base.description ?? fallback.description ?? "",
        Name: base.Name ?? fallback.Name ?? "",
        categories: base.categories ?? fallback.categories ?? "",
        image: base.image ?? fallback.image ?? "",
      } as RowData;
    }
    if (skuimgResult.length === 0) {
      throw new Error("No data to process.");
    }
    if (!upload) {
      throw new Error("No file uploaded");
    }
    const data = await upload.arrayBuffer();
    const textData = new TextDecoder().decode(data);
    const wb = read(textData, { type: "string" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const wcProductData = utils.sheet_to_json<Record<string, string>>(sheet);
    attributeResult = skuimgResult.map((item) => {
      const matchedRow = wcProductData.find((row) => row["Name"] === item.Name);
      let attribute = "";

      if (matchedRow && matchedRow["Attribute 1 name"]?.trim()) {
        attribute = matchedRow["Attribute 1 name"].trim();
      } else if (item.type === "variable") {
        attribute = `65smart${Math.floor(100000 + Math.random() * 900000)}`;
      }

      return { ...item, attribute };
    });

    const grouped = new Map<string, RowData[]>();
    const groupedArray: RowData[][] = [];
    attributeResult.forEach((item) => {
      const groupKey = item.lazada_product_id ?? `__ungrouped__${item.SKU}`;
      if (!grouped.has(groupKey)) grouped.set(groupKey, []);
      grouped.get(groupKey)!.push(item);
    });
    grouped.forEach((items, groupKey) => {
      const groupRows: RowData[] = [];
      const variations = items.filter((i) => i.type === "variation");
      const variables = items.filter((i) => i.type === "variable");
      const simples = items.filter((i) => i.type === "simple");
      if (variations.length > 0) {
        const existingVariableRow = wcProductData.find(
          (row) => row["Meta: lazada_product_id"] === groupKey && row["Type"] === "variable"
        );
        const variableSku = existingVariableRow
          ? existingVariableRow["SKU"]
          : `65smarttools-${variations[0]?.SKU ?? "unknown"}`;
        const variable = fillDefaultFields({
          ...variables[0],
          type: "variable",
          SKU: variableSku,
          attribute: variables[0]?.attribute ?? variations[0]?.attribute ?? "",
          attributeValue: "",
        }, variations[0]);
        const uniqueValues = Array.from(
          new Set(
            variations.map((v) => v.attributeValue?.trim()).filter(Boolean)
          )
        );
        variable.attributeValue = uniqueValues.join(",");
        const attrName = variable.attribute?.trim() ?? "";
        const swatch = {
          [attrName]: {
            name: attrName,
            type: attrName === "select" ? "select" : "image",
            terms: Object.fromEntries(
              uniqueValues.map((value) => {
                const matchedVariation = variations.find(
                  (v) => v.attributeValue?.trim() === value
                );
                return [
                  value,
                  {
                    name: value,
                    color: "",
                    image: matchedVariation?.image || false,
                    show_tooltip: "",
                    tooltip_text: "",
                    tooltip_image: "",
                    image_size: "38448",
                  },
                ];
              })
            ),
          },
        };
        variable.json = JSON.stringify(swatch);
        variable.swatchesAttributes = JSON.stringify(swatch);
        groupRows.push(variable);
        variations.forEach((v) => {
          groupRows.push({
            ...v,
            parent: variableSku,
            attribute: variable.attribute,
            image: v.image,
          });
        });
      }
      simples.forEach((s) => groupRows.push(s));
      groupedArray.push(groupRows);
    });
    const finalGrouped = groupedArray.reverse().flat();
    const exportRows = finalGrouped.map((item) => ({
      ID: item.ID,
      Type: item.type,
      SKU: item.SKU,
      "GTIN, UPC, EAN, or ISBN": "",
      Name: item.Name,
      Published: 1,
      "Is featured?": 1,
      "Visibility in catalog": "visible",
      "Short description": item.shortDescription,
      Description: item.description,
      "Date sale price starts": "",
      "Date sale price ends": "",
      "Tax status": "taxable",
      "Tax class": "",
      "In stock?": "",
      Stock: item.stock,
      "Low stock amount": "",
      "Backorders allowed?": 0,
      "Sold individually?": 0,
      "Weight (kg)": item.weight,
      "Length (cm)": item.length,
      "Width (cm)": item.width,
      "Height (cm)": item.height,
      "Allow customer reviews?": 0,
      "Purchase note": "",
      "Sale price": item.salePrice,
      "Regular price": item.regularPrice,
      Categories: item.categories,
      Tags: "",
      "Shipping class": "",
      Images: item.image,
      "Download limit": "",
      "Download expiry days": "",
      Parent: item.parent,
      "Grouped products": "",
      Upsells: "",
      "Cross-sells": "",
      "External URL": "",
      "Button text": "",
      Position: "",
      "Swatches Attributes": item.swatchesAttributes,
      Brand: item.brand,
      "Attribute 1 name": item.attribute,
      "Attribute 1 value(s)": item.attributeValue,
      "Attribute 1 global": 1,
      "Meta: is_installment_variable_attributes": item.installment_variable,
      "Meta: rtwpvg_images": item.rtwpvg_images,
      "Meta: lazada_product_id": item.lazada_product_id,
    }));
    const newSheet = utils.json_to_sheet(exportRows);
    const newWb = utils.book_new();
    utils.book_append_sheet(newWb, newSheet, "Result");
    const csv = utils.sheet_to_csv(newSheet);
    return new Response(Buffer.from(csv), {
      headers: {
        "Content-Type": "text/csv",
        "Content-Disposition": 'attachment; filename="update_products_woocommerce.csv"',
      },
    });
  },
  {
    body: t.Object({ upload: t.File() }),
  }
).listen(process.env.PORT || 3001, () =>
  console.log("✅ Elysia server running at http://localhost:3001")
);