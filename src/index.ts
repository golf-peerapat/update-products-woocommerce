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
  tags?: string;
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
let finalResult: RowData[] = [];
type JsonType = Record<string, AttributeData>;
const app = new Elysia();
let jsonType: JsonType[] = [];
dotenv.config();

app.onRequest(({ set }) => {
  set.headers["Access-Control-Allow-Origin"] = "*";
  set.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS";
  set.headers["Access-Control-Allow-Headers"] = "Content-Type";
});

app.post(
  "/wc-product-export",
  async ({ body: { upload } }) => {
    if (skuimgResult.length === 0) throw new Error("No data to process.");
    if (!upload) throw new Error("No file uploaded");

    const data = await upload.arrayBuffer();
    const textData = new TextDecoder().decode(data);
    const wb = read(textData, { type: "string" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const wcProductData = utils.sheet_to_json<Record<string, string>>(sheet);

    // 👇 Logic เดิมทำงานตามปกติ (ไม่ต้องแก้)
    attributeResult = skuimgResult.map((item) => {
      const skuName = item.Name;
      const matchedRow = wcProductData.find((row) => row["Name"] === skuName);
      let attribute = "";
      if (matchedRow) {
        attribute = matchedRow["Attribute 1 name"];
      } else if (item.type === "variable") {
        attribute = `65smart${Math.floor(100000 + Math.random() * 900000)}`;
      }
      return {
        ...item,
        attribute,
      };
    });

    attributeResult = attributeResult.map((item) => {
      if (item.parent?.startsWith("id:")) {
        const parentId = item.parent.replace("id:", "").trim();
        const parentItem = attributeResult.find(
          (parent) => parent.ID === parentId
        );
        if (parentItem) {
          return {
            ...item,
            attribute: parentItem.attribute,
          };
        }
      }
      return item;
    });

    jsonType = attributeResult
      .filter((item) => item.attributeValue && item.type === "variation")
      .map((item) => ({
        [`${item.attributeValue}`]: {
          name: item.attributeValue || "",
          color: "",
          image: item.image || "",
          show_tooltip: "",
          tooltip_text: "",
          tooltip_image: "",
          image_size: "38448",
        },
      }));

    jsonResult = attributeResult.map((item) => {
      const jsonData = jsonType.find((data) => data[item.attributeValue || ""]);
      return {
        ...item,
        json: jsonData ? JSON.stringify(jsonData) : "",
      };
    });

    const combinedJson: Record<string, any> = {};
    jsonResult.forEach((parentItem) => {
      if (parentItem.attribute && parentItem.ID) {
        if (!combinedJson[parentItem.attribute]) {
          combinedJson[parentItem.attribute] = {
            name: parentItem.attribute,
            type: "image",
            terms: {},
          };
        }
        const childItems = jsonResult.filter((child) => {
          if (child.parent) {
            const childParentId = String(child.parent)
              .replace("id:", "")
              .trim();
            const parentId = String(parentItem.ID);
            return childParentId === parentId;
          }
          return false;
        });
        childItems.forEach((child) => {
          if (child.json) {
            try {
              const jsonData = JSON.parse(child.json);
              combinedJson[parentItem.attribute as string].terms = {
                ...combinedJson[parentItem.attribute as string].terms,
                ...jsonData,
              };
            } catch (error) {
              console.error("❌ Error parsing JSON:", error);
            }
          }
        });
      }
    });

    jsonResult = jsonResult.map((item) => {
      const swatchesAttributes = combinedJson[item.attribute as string] || null;
      return {
        ...item,
        swatchesAttributes:
          item.type === "variation"
            ? ""
            : swatchesAttributes
            ? JSON.stringify(swatchesAttributes)
            : "",
      };
    });
    const newSheet = utils.json_to_sheet(
      jsonResult.map(
        ({
          ID,
          type,
          SKU,
          Name,
          shortDescription,
          description,
          stock,
          weight,
          length,
          width,
          height,
          salePrice,
          regularPrice,
          categories,
          tags,
          image,
          parent,
          swatchesAttributes,
          brand,
          attribute,
          attributeValue,
          installment_variable,
          rtwpvg_images,
        }) => ({
          ID: ID,
          Type: type,
          SKU: SKU,
          "GTIN, UPC, EAN, or ISBN": "",
          Name: Name,
          Published: 1,
          "Is featured?": 1,
          "Visibility in catalog": "visible",
          "Short description": shortDescription,
          Description: description,
          "Date sale price starts": "",
          "Date sale price ends": "",
          "Tax status": "taxable",
          "Tax class": "",
          "In stock?": "",
          Stock: stock,
          "Low stock amount": "",
          "Backorders allowed?": 0,
          "Sold individually?": 0,
          "Weight (kg)": weight,
          "Length (cm)": length,
          "Width (cm)": width,
          "Height (cm)": height,
          "Allow customer reviews?": 0,
          "Purchase note": "",
          "Sale price": salePrice,
          "Regular price": regularPrice,
          Categories: categories,
          Tags: tags,
          "Shipping class": "",
          Images: image,
          "Download limit": "",
          "Download expiry days": "",
          Parent: parent,
          "Grouped products": "",
          Upsells: "",
          "Cross-sells": "",
          "External URL": "",
          "Button text": "",
          Position: "",
          "Swatches Attributes": swatchesAttributes,
          Brand: brand,
          "Attribute 1 name": attribute,
          "Attribute 1 value(s)": attributeValue,
          "Attribute 1 global": 1,
          "Meta: is_installment_variable_attributes": installment_variable,
          "Meta: rtwpvg_images": rtwpvg_images,
        })
      )
    );
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
);

app.post(
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
      jsonData.forEach((row) => {
        const productId = row[0] as string | null;
        const nameProduct = row[2] as string | null;
        const skuProduct = row[15] as string | null;
        const variationsCombo = row[16] as string | null;
        const imageData = row.slice(7, 15).filter(Boolean).join(",");
        if (productId) {
          if (seen.has(productId)) {
            const existingAttribute = seen.get(productId)?.attributeValue || "";
            const combinedAttribute = [existingAttribute, variationsCombo]
              .filter(Boolean)
              .join(",");
            seen.set(productId, {
              ID: productId,
              type: "variable",
              SKU: "",
              Name: nameProduct ?? "",
              tags: "suggestion_item",
              image: seen.get(productId)?.image ?? imageData,
              attributeValue: combinedAttribute,
            });
          } else {
            seen.set(productId, {
              ID: productId,
              type: "simple",
              SKU: skuProduct ?? "",
              Name: nameProduct ?? "",
              tags: "suggestion_item",
              image: imageData,
              attributeValue: "",
            });
          }
        }
        if (variationsCombo) {
          variations.push({
            ID: "",
            type: "variation",
            SKU: skuProduct ?? "",
            Name: `${nameProduct ?? ""} - ${variationsCombo}`,
            image: row[7] ?? "",
            parent: `id:${productId}`,
            attributeValue: variationsCombo,
            installment_variable: "yes",
            rtwpvg_images: imageData,
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
);

app.post(
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
      const newImage = imageMap.get(item.ID);
      return {
        ...item,
        shortDescription: shortDescriptionMap.get(item.ID) || "",
        description: descriptionMap.get(item.ID) || "",
        image: newImage || item.image,
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
);

app.post(
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
        .map((value) => Number(value))
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
);

app.post(
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
);

app.post(
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
    skuimgResult = skuimgResult.map((item) => ({
      ...item,
      categories: tabNameMap.get(item.ID) || "",
      brand: brandMap.get(item.ID) || "",
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
);

app.listen(process.env.PORT || 3001, () =>
  console.log("✅ Elysia server running at http://localhost:3001")
);