import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Home, Wine, Martini, Beer, CupSoda, Snowflake,
  Search, ShoppingCart, MessageCircle, Trash2, Camera,
  Minus, Plus, Package, Truck, Percent, Headphones, ChevronRight,
  LoaderCircle, AlertCircle
, Banknote, FileText, CreditCard, ArrowLeftRight} from "lucide-react";
import "./styles.css";

const JSON_URL = "/data/products-list.json";
const BEST_SELLER_FILES = {
  DESTACADOS: "/data/best_sellers_destacados.xlsx",
  VINOS: "/data/best_sellers_wines.xlsx",
  ESPUMANTES: "/data/best_sellers_espumantes.xlsx",
  LICORES: "/data/best_sellers_licores.xlsx",
  CERVEZAS: "/data/best_sellers_cervezas.xlsx",
  BEBIDAS: "/data/best_sellers_bebidas.xlsx",
  CONGELADOS: "/data/best_sellers_congelados.xlsx",
};

const categoryBannerConfig = {
  VINOS: { title: "Vinos", subtitle: "Descubre los vinos más elegidos de El Palmar", eyebrow: "SELECCIÓN EL PALMAR", icon: "🍷", image: "/category-banners/vinos.png" },
  ESPUMANTES: { title: "Espumantes", subtitle: "Descubre los espumantes más elegidos de El Palmar", eyebrow: "SELECCIÓN EL PALMAR", icon: "🥂", image: "/category-banners/espumantes.png" },
  LICORES: { title: "Licores", subtitle: "Descubre los licores más elegidos de El Palmar", eyebrow: "SELECCIÓN EL PALMAR", icon: "🍸", image: "/category-banners/licores.png" },
  CERVEZAS: { title: "Cervezas", subtitle: "Descubre las cervezas más elegidas de El Palmar", eyebrow: "SELECCIÓN EL PALMAR", icon: "🍺", image: "/category-banners/cervezas.png" },
  BEBIDAS: { title: "Bebidas", subtitle: "Descubre las bebidas más elegidas de El Palmar", eyebrow: "SELECCIÓN EL PALMAR", icon: "🥤", image: "/category-banners/bebidas.png" },
  CONGELADOS: { title: "Congelados", subtitle: "Descubre los productos congelados más elegidos de El Palmar", eyebrow: "SELECCIÓN EL PALMAR", icon: "❄️", image: "/category-banners/congelados.png" },
};

const categoryGroups = [
  { key: "ALL", label: "Inicio", icon: Home },
  { key: "DESTACADOS", label: "Destacados", icon: Percent },
  { key: "VINOS", label: "Vinos", icon: Wine },
  { key: "ESPUMANTES", label: "Espumantes", icon: Wine },
  { key: "LICORES", label: "Licores", icon: Martini },
  { key: "CERVEZAS", label: "Cervezas", icon: Beer },
  { key: "BEBIDAS", label: "Bebidas", icon: CupSoda },
  { key: "CONGELADOS", label: "Congelados", icon: Snowflake },
];

const groupMap = {
  VINOS: ["VINOS"],
  ESPUMANTES: ["ESPUMANTES"],
  LICORES: ["LICORES"],
  CERVEZAS: ["CERVEZAS"],
  BEBIDAS: ["BEBIDAS NO ALCOHOLICAS"],
  CONGELADOS: ["CONGELADOS"],
  DESTACADOS: ["DESTACADOS"],
};

const money = (value) => new Intl.NumberFormat("es-CL", {
  style: "currency", currency: "CLP", maximumFractionDigits: 0
}).format(Number(value) || 0);

function normalizeProduct(row, index) {
  const stock = Number(row.stock);
  return {
    id: String(row.id ?? index),
    barcode: String(row.codigoBarras ?? "").replace(/\.0$/, ""),
    category: String(row.categoria ?? "").trim(),
    family: String(row.familia ?? "").trim(),
    name: String(row.nombre ?? "").trim(),
    price: Number(row.precio) || 0,
    stock: Number.isFinite(stock) ? stock : 0,
    image: typeof row.imagen === "string" && row.imagen.trim() ? row.imagen.trim() : null,
  };
}

async function loadProductsFromJson() {
  const response = await fetch(`${JSON_URL}?v=${Date.now()}`);
  if (!response.ok) throw new Error(`No se pudo cargar ${JSON_URL}`);

  const payload = await response.json();
  const rows = Array.isArray(payload)
    ? payload
    : payload?.productos || payload?.products || payload?.data || [];

  return rows
    .map(normalizeProduct)
    .filter((p) => p.name && p.category);
}

function normalizeBarcode(value) {
  return String(value ?? "")
    .replace(/\.0$/, "")
    .replace(/\s+/g, "")
    .replace(/-/g, "");
}

async function loadBestSellersFile(url) {
  const response = await fetch(`${url}?v=${Date.now()}`);
  if (!response.ok) throw new Error(`No se pudo cargar ${url}`);

  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" });

  return rows
    .map((row) => ({
      barcode: normalizeBarcode(row["Código Barra"]),
      description: String(row["Descripción"] ?? "").trim(),
      category: String(row["Categoría"] ?? "").trim(),
      family: String(row["Familia"] ?? "").trim(),
      quantitySold: Number(row["Cantidad Vendida"]) || 0,
      revenue: Number(row["Ingresos Generados"]) || 0,
    }))
    .filter((row) => row.barcode && row.quantitySold > 0);
}

function App() {
  const [products, setProducts] = useState([]);
  const [selectedCategory, setSelectedCategory] = useState("ALL");
  const [selectedFamily, setSelectedFamily] = useState("ALL");
  const [catalogSort, setCatalogSort] = useState("DEFAULT");
  const [search, setSearch] = useState("");
  const [searchFocused, setSearchFocused] = useState(false);
  const [activeSearchSuggestion, setActiveSearchSuggestion] = useState(-1);
  const [searchSelectedProductId, setSearchSelectedProductId] = useState(null);
  const [familyFiltersExpanded, setFamilyFiltersExpanded] = useState(true);
  const [currentPage, setCurrentPage] = useState(1);
  const [cart, setCart] = useState([]);
  const [cartOpen, setCartOpen] = useState(false);
  const [infoOpen, setInfoOpen] = useState(true);
  const [imageProduct, setImageProduct] = useState(null);
  const [imageZoom, setImageZoom] = useState(1);
  const cartRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [promoPage, setPromoPage] = useState(0);
  const [promoTimerReset, setPromoTimerReset] = useState(0);
  const [promosPerPage, setPromosPerPage] = useState(4);
  const [bestSellerRowsByCategory, setBestSellerRowsByCategory] = useState({});

  useEffect(() => {
    const updatePromosPerPage = () => {
      setPromosPerPage(window.innerWidth <= 850 ? 2 : 4);
    };

    updatePromosPerPage();
    window.addEventListener("resize", updatePromosPerPage);
    return () => window.removeEventListener("resize", updatePromosPerPage);
  }, []);

  useEffect(() => {
    if (!imageProduct) return;

    const handleKeyDown = (event) => {
      if (event.key === "Escape") {
        setImageProduct(null);
        setImageZoom(1);
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [imageProduct]);

  useEffect(() => {
    if (!imageProduct) return;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      document.body.style.overflow = previousOverflow;
    };
  }, [imageProduct]);

  useEffect(() => {
    if (cartOpen) {
      requestAnimationFrame(() => {
        cartRef.current?.scrollTo({
          top: 0,
          behavior: "auto"
        });
      });
    }
  }, [cartOpen]);

  useEffect(() => {
    loadProductsFromJson()
      .then(setProducts)
      .catch((err) => setError(err.message))
      .finally(() => setLoading(false));
  }, []);

  // Cuando cambia la categoría, vuelve al inicio de la zona principal para
  // mostrar primero el banner de esa categoría y luego sus contenidos.
  useEffect(() => {
    const main = document.querySelector(".main");
    main?.scrollTo({ top: 0, behavior: "auto" });
  }, [selectedCategory]);

  useEffect(() => {
    let cancelled = false;

    Promise.all(
      Object.entries(BEST_SELLER_FILES).map(async ([category, url]) => {
        try {
          const rows = await loadBestSellersFile(url);
          return [category, rows];
        } catch {
          return [category, []];
        }
      })
    ).then((entries) => {
      if (!cancelled) setBestSellerRowsByCategory(Object.fromEntries(entries));
    });

    return () => { cancelled = true; };
  }, []);

  const featuredProducts = useMemo(
    () => [...products]
      .filter(p => p.category === "DESTACADOS" && p.stock > 0)
      .sort((a, b) => a.name.localeCompare(b.name, "es", { sensitivity: "base" })),
    [products]
  );

  const promoTotalPages = Math.ceil(featuredProducts.length / promosPerPage);
  const visibleFeaturedProducts = featuredProducts.slice(
    promoPage * promosPerPage,
    (promoPage + 1) * promosPerPage
  );

  useEffect(() => {
    setPromoPage(page => Math.min(page, Math.max(0, promoTotalPages - 1)));
  }, [promosPerPage, promoTotalPages]);

  useEffect(() => {
    if (promoTotalPages <= 1) {
      setPromoPage(0);
      return;
    }

    const timer = setInterval(() => {
      setPromoPage(page => (page + 1) % promoTotalPages);
    }, 6000);

    return () => clearInterval(timer);
  }, [promoTotalPages, promoTimerReset]);

  function changePromoPage(delta) {
    setPromoPage(page => (page + delta + promoTotalPages) % promoTotalPages);
    setPromoTimerReset(value => value + 1);
  }

  function addFromImageModal() {
    if (!imageProduct || imageProduct.stock <= 0) return;
    addToCart(imageProduct);
    setImageProduct(null);
    setImageZoom(1);
  }

  const searchSuggestions = useMemo(() => {
    const query = search.toLowerCase().trim();

    if (!query) return [];

    const seen = new Set();

    const catalogCategories = new Set(
      Object.values(groupMap).flat()
    );

    return products
      .filter(p => {
        // El buscador solo debe sugerir productos pertenecientes a las
        // categorías que realmente forman parte del catálogo.
        if (p.stock <= 0 || !catalogCategories.has(p.category)) return false;

        return [p.name, p.family, p.category, p.barcode]
          .some(v => String(v ?? "").toLowerCase().includes(query));
      })
      .sort((a, b) => {
        const aName = a.name.toLowerCase();
        const bName = b.name.toLowerCase();

        const aStarts = aName.startsWith(query) ? 0 : 1;
        const bStarts = bName.startsWith(query) ? 0 : 1;

        if (aStarts !== bStarts) return aStarts - bStarts;
        return aName.localeCompare(bName, "es", { sensitivity: "base" });
      })
      .filter(p => {
        const key = p.name.toLowerCase();

        if (seen.has(key)) return false;

        seen.add(key);
        return true;
      })
      .slice(0, 6);
  }, [products, search]);

  const showSearchSuggestions =
    searchFocused &&
    search.trim().length > 0 &&
    searchSuggestions.length > 0;

  function commitSearchResult(product = null) {
    const query = search.toLowerCase().trim();

    const target =
      product ||
      products.find(p =>
        p.stock > 0 &&
        Object.values(groupMap).flat().includes(p.category) &&
        p.name.toLowerCase() === query
      ) ||
      searchSuggestions[0];

    if (!target) return;

    // Guarda el producto que originó la búsqueda para destacarlo dentro de su familia.
    setSearchSelectedProductId(target.id);

    const category = Object.entries(groupMap).find(([, categories]) =>
      categories.includes(target.category)
    )?.[0];

    if (category) {
      setSelectedCategory(category);
    }

    // La búsqueda identifica el producto, pero una vez encontrado
    // el filtro pasa a categoría + familia.
    setSelectedFamily(target.family || "ALL");
    setFamilyFiltersExpanded(false);
    setCurrentPage(1);
    setSearch("");
    setSearchFocused(false);
    setActiveSearchSuggestion(-1);

    setTimeout(() => {
      const targetCard = document.getElementById(`product-${target.id}`);

      if (targetCard) {
        targetCard.scrollIntoView({
          behavior: "smooth",
          block: "center"
        });
      } else {
        document.getElementById("products-section")?.scrollIntoView({
          behavior: "smooth",
          block: "start"
        });
      }
    }, 120);
  }

  function scrollToSearchResults() {
    commitSearchResult();
  }

  function selectSearchSuggestion(product) {
    commitSearchResult(product);
  }

  function highlightSearchMatch(text) {
    const query = search.trim();

    if (!query) return text;

    const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const parts = String(text).split(new RegExp(`(${escapedQuery})`, "ig"));

    return parts.map((part, index) =>
      part.toLowerCase() === query.toLowerCase()
        ? <strong key={index}>{part}</strong>
        : <React.Fragment key={index}>{part}</React.Fragment>
    );
  }

  const searchCategory = useMemo(() => {
    const query = search.toLowerCase().trim();

    if (!query) return null;

    // Si la búsqueda coincide exactamente con una categoría, usamos esa categoría.
    const exactCategory = categoryGroups.find(
      ({ key, label }) =>
        key !== "ALL" &&
        label.toLowerCase() === query
    )?.key;

    if (exactCategory) return exactCategory;

    // Desde Inicio, si todos los resultados de la búsqueda pertenecen
    // a una sola categoría, la mostramos automáticamente como categoría activa.
    const matchedCategories = new Set(
      products
        .filter(p => {
          if (p.stock <= 0 || p.category === "PROMOCIONES") return false;
          return [p.name, p.family, p.category, p.barcode]
            .some(v => String(v ?? "").toLowerCase().includes(query));
        })
        .map(p => p.category)
        .filter(category =>
          categoryGroups.some(group =>
            group.key !== "ALL" && groupMap[group.key]?.includes(category)
          )
        )
    );

    return matchedCategories.size === 1
      ? [...matchedCategories][0] === "BEBIDAS NO ALCOHOLICAS"
        ? "BEBIDAS"
        : [...matchedCategories][0]
      : null;
  }, [search, products]);

  const categoryForDisplay = searchCategory || selectedCategory;
  const showHomePromos = categoryForDisplay === "ALL";

  const currentBestSellerRows = bestSellerRowsByCategory[categoryForDisplay] || [];

  const availableFamilies = useMemo(() => {
    const categoryForFamilies = searchCategory || selectedCategory;
    const isSearchingFromInicio = Boolean(search.trim()) && categoryForFamilies === "ALL";

    const cats = categoryForFamilies === "ALL"
      ? (isSearchingFromInicio
          ? Object.values(groupMap).flat()
          : Object.entries(groupMap)
              .filter(([key]) => key !== "DESTACADOS")
              .flatMap(([, categories]) => categories))
      : groupMap[categoryForFamilies];

    // Después de una búsqueda, mostramos únicamente la familia
    // encontrada. "Todas las familias" permite volver a expandir la lista.
    if (!familyFiltersExpanded && selectedFamily !== "ALL") {
      return [selectedFamily];
    }

    const families = products
      .filter(p => cats?.includes(p.category))
      .map(p => p.family)
      .filter(Boolean);

    return [...new Set(families)].sort((a, b) => a.localeCompare(b, "es"));
  }, [
    products,
    selectedCategory,
    searchCategory,
    selectedFamily,
    familyFiltersExpanded
  ]);

  // Ventas por código de barras para ordenar el catálogo por popularidad.
  // El reporte de ventas se mantiene separado del catálogo y se cruza por código.
  const salesByBarcode = useMemo(() => {
    const sales = new Map();
    currentBestSellerRows.forEach((row) => {
      const barcode = normalizeBarcode(row.barcode);
      if (!barcode) return;
      sales.set(barcode, (sales.get(barcode) || 0) + (Number(row.quantitySold) || 0));
    });
    return sales;
  }, [currentBestSellerRows]);

  // Destacados también funciona como una categoría del catálogo.
  // La sección superior de destacados se mantiene independiente.
  const filteredProducts = useMemo(() => {
    const query = search.toLowerCase().trim();
    const categoryForSearch = searchCategory || selectedCategory;
    // En "Inicio" solo mostramos productos de las categorías
    // que existen en la barra principal. No se muestran categorías
    // internas como PASTELERIA, VARIOS, CECINAS Y LACTEOS, etc.
    const visibleCategories = Object.entries(groupMap)
      .filter(([key]) => key !== "DESTACADOS")
      .flatMap(([, categories]) => categories);
    const isSearchingFromInicio = Boolean(query) && categoryForSearch === "ALL";
    const cats = categoryForSearch === "ALL"
      ? (isSearchingFromInicio ? Object.values(groupMap).flat() : visibleCategories)
      : groupMap[categoryForSearch];
    return products
      .filter((p) => {
        const matchesCategory = !cats || cats.includes(p.category);
        const matchesStock = p.stock > 0;
        const matchesFamily = selectedFamily === "ALL" || p.family === selectedFamily;
        // Si el usuario selecciona una familia después de buscar,
        // la familia pasa a ser el filtro principal y el texto del buscador
        // deja de limitar los resultados.
        const matchesSearch =
          selectedFamily !== "ALL" ||
          !query ||
          [p.name, p.family, p.category, p.barcode]
            .some(v => v.toLowerCase().includes(query));
        const showDestacados = categoryForSearch === "DESTACADOS" || isSearchingFromInicio;
        return matchesCategory &&
          matchesStock &&
          matchesFamily &&
          matchesSearch &&
          p.category !== "PROMOCIONES" &&
          (showDestacados || p.category !== "DESTACADOS");
      })
      .sort((a, b) => {
        if (searchSelectedProductId) {
          if (a.id === searchSelectedProductId && b.id !== searchSelectedProductId) return -1;
          if (b.id === searchSelectedProductId && a.id !== searchSelectedProductId) return 1;
        }

        if (catalogSort === "BEST_SELLERS") {
          const salesA = salesByBarcode.get(normalizeBarcode(a.barcode)) || 0;
          const salesB = salesByBarcode.get(normalizeBarcode(b.barcode)) || 0;
          if (salesA !== salesB) return salesB - salesA;
        }

        if (catalogSort === "PRICE_ASC" && a.price !== b.price) {
          return a.price - b.price;
        }

        if (catalogSort === "PRICE_DESC" && a.price !== b.price) {
          return b.price - a.price;
        }

        if (catalogSort === "NAME") {
          return a.name.localeCompare(b.name, "es", { sensitivity: "base" });
        }

        return 0;
      });
  }, [products, selectedCategory, selectedFamily, search, searchCategory, searchSelectedProductId, catalogSort, salesByBarcode]);

  const cartUnits = cart.reduce((sum, item) => sum + item.quantity, 0);

  // Descuento de hielo por formato:
  // - Hielo de 2 kg: desde 30 bolsas -> $100 menos por bolsa.
  // - Hielo de 1 kg: desde 50 bolsas -> $100 menos por bolsa.
  // Cada formato se calcula por separado.
  const ice2kgUnits = cart
    .filter(item => {
      const text = `${item.name} ${item.family}`.toLowerCase();
      return /hielo/.test(text) && /\b2\s*(?:kg|kl|kilos?)\b/i.test(text);
    })
    .reduce((sum, item) => sum + item.quantity, 0);

  const ice1kgUnits = cart
    .filter(item => {
      const text = `${item.name} ${item.family}`.toLowerCase();
      return /hielo/.test(text) && /\b1\s*(?:kg|kl|kilos?)\b/i.test(text);
    })
    .reduce((sum, item) => sum + item.quantity, 0);

  const cartSubtotal = cart.reduce(
    (sum, item) => sum + item.price * item.quantity,
    0
  );

  const iceDiscount =
    (ice2kgUnits >= 30 ? ice2kgUnits * 100 : 0) +
    (ice1kgUnits >= 50 ? ice1kgUnits * 100 : 0);

  const cartTotal = cartSubtotal - iceDiscount;

  const PRODUCTS_PER_PAGE = 24;
  const totalPages = Math.ceil(filteredProducts.length / PRODUCTS_PER_PAGE);
  const paginatedProducts = filteredProducts.slice(
    (currentPage - 1) * PRODUCTS_PER_PAGE,
    currentPage * PRODUCTS_PER_PAGE
  );

  function selectCategory(key) {
    setSelectedCategory(key);
    setSearchSelectedProductId(null);
    setSelectedFamily("ALL");
    setCatalogSort("DEFAULT");
    setFamilyFiltersExpanded(true);
    setSearch("");
    setCurrentPage(1);
  }

  // Cada cambio de categoría inicia el catálogo con filtros limpios.
  // Esto evita que un orden/familia seleccionado en una categoría
  // se arrastre a la siguiente categoría.
  useEffect(() => {
    setSelectedFamily("ALL");
    setCatalogSort("DEFAULT");
    setFamilyFiltersExpanded(true);
    setCurrentPage(1);
  }, [selectedCategory]);


  function selectFamily(family) {
    setSelectedFamily(family);
    setSearchSelectedProductId(null);
    setFamilyFiltersExpanded(true);
    setSearch("");
    setCurrentPage(1);
  }

  function goHome() {
    setSelectedCategory("ALL");
    setSearchSelectedProductId(null);
    setCurrentPage(1);
    setSelectedFamily("ALL");
    setFamilyFiltersExpanded(true);
    setSearch("");
    document.querySelector(".main")?.scrollTo({
      top: 0,
      behavior: "smooth"
    });
  }

  function scrollToProducts() {
    setTimeout(() => {
      document.getElementById("products-section")?.scrollIntoView({
        behavior: "smooth",
        block: "start"
      });
    }, 50);
  }

  function addToCart(product) {
    if (product.stock <= 0) return;

    setCartOpen(true);

    setCart(prev => {
      const existing = prev.find(i => i.id === product.id);
      if (existing) {
        if (existing.quantity >= Math.floor(product.stock)) return prev;
        return prev.map(i => i.id === product.id ? { ...i, quantity: i.quantity + 1 } : i);
      }
      return [...prev, { ...product, quantity: 1 }];
    });
  }

  function changeQty(id, delta) {
    setCart(prev => prev.flatMap(item => {
      if (item.id !== id) return [item];
      const next = item.quantity + delta;
      if (next <= 0) return [];
      if (next > Math.floor(item.stock)) return [item];
      return [{ ...item, quantity: next }];
    }));
  }

  function setQty(id, value) {
    const max = Math.max(1, Math.floor(
      cart.find(item => item.id === id)?.stock || 1
    ));
    const next = Math.min(Math.max(1, Number(value) || 1), max);

    setCart(prev =>
      prev.map(item =>
        item.id === id ? { ...item, quantity: next } : item
      )
    );
  }

  function openWhatsAppContact() {
    const whatsappNumber = "988137633";
    window.open(`https://wa.me/${whatsappNumber}`, "_blank");
  }

  function sendWhatsApp() {
    if (!cart.length) return;

    const whatsappNumber = "988137633";
    const lines = cart.map(
      i => `• ${i.name} x${i.quantity} — ${money(i.price * i.quantity)}`
    );

    const discountLine = iceDiscount > 0
      ? `\u{1F3AF} Descuento por volumen en hielo: -${money(iceDiscount)}\n`
      : "";

    const message = `\u{1F3EA} Hola Distribuidora El Palmar

\u{1F6D2} Quisiera cotizar el siguiente pedido:

${lines.join("\n")}

\u{1F4B5} Subtotal: ${money(cartSubtotal)}
${discountLine}\u{1F4B0} Total con descuento: ${money(cartTotal)}

\u{1F4CB} Quedo atento/a para confirmar el pedido.
¡Muchas gracias! \u{1F60A}`;

    window.open(
      `https://wa.me/${whatsappNumber}?text=${encodeURIComponent(message)}`,
      "_blank"
    );
  }

  const currentCategoryConfig = categoryBannerConfig[categoryForDisplay];

  const hero = categoryForDisplay === "ALL" ? (
    <section className="hero">
      <img src="/el-palmar-banner.png" alt="Distribuidora El Palmar" />
    </section>
  ) : (
    <section className="hero category-hero" data-category-hero="true">
      <div className="category-banner category-banner-full">
        {currentCategoryConfig?.image ? (
          <img
            src={currentCategoryConfig.image}
            alt={`Banner de ${currentCategoryConfig.title} de Distribuidora El Palmar`}
            onError={(event) => {
              event.currentTarget.style.display = "none";
              event.currentTarget.parentElement?.classList.add("category-banner-placeholder");
            }}
          />
        ) : null}
        <div className="category-banner-content">
          <span>{currentCategoryConfig?.eyebrow || "SELECCIÓN EL PALMAR"}</span>
          <strong>{currentCategoryConfig?.icon || "✦"} {currentCategoryConfig?.title || categoryForDisplay}</strong>
          <p>{currentCategoryConfig?.subtitle || "Explora nuestros productos y encuentra lo que buscas."}</p>
        </div>
      </div>
    </section>
  );

  return (
    <div className="app">
      <aside className={`sidebar ${cartOpen ? "sidebar-cart-open" : ""}`}>
        <div className="side-title-row">
          <div className="side-title">CATEGORÍAS</div>
          <div className="category-scroll-hint" aria-hidden="true">
            <span>Desliza para ver más categorías</span>
            <ChevronRight size={16} />
          </div>
        </div>
        <div className="side-line" />
        <nav>
          <button
            className={`side-item ${(searchCategory || selectedCategory) === "ALL" ? "active" : ""}`}
            onClick={goHome}
          >
            <Home size={21} /><span>Inicio</span>
          </button>

          {categoryGroups
            .filter(({ key }) => key !== "ALL")
            .map(({ key, label, icon: Icon }) => (
              <button
                key={key}
                className={`side-item ${(searchCategory || selectedCategory) === key ? "active" : ""}`}
                onClick={() => selectCategory(key)}
              >
                <Icon size={21} /><span>{label}</span>
              </button>
            ))}
        </nav>
        <div className="side-tag">Tu distribuidora<br />de confianza</div>
      </aside>

      <main className="main">
        {hero}
        <div className="search-wrap">
          <Search size={27} className="search-main-icon" />

          <div className="search-box">
            <input
              value={search}
              onFocus={() => {
                setSearchFocused(true);
                setActiveSearchSuggestion(-1);

                if (search.trim()) {
                  setSelectedFamily("ALL");
                  setFamilyFiltersExpanded(true);
                  setSelectedCategory("ALL");
                  setCurrentPage(1);
                }
              }}
              onBlur={() => {
                setTimeout(() => setSearchFocused(false), 120);
              }}
              onChange={e => {
                const value = e.target.value;

                setSearch(value);
                setActiveSearchSuggestion(-1);
                if (value.trim()) setSearchSelectedProductId(null);

                // Si el usuario empieza una nueva búsqueda después de haber
                // seleccionado una familia, liberamos ese filtro para que
                // la nueva búsqueda sea completamente independiente.
                if (value.trim()) {
                  setSelectedFamily("ALL");
                  setFamilyFiltersExpanded(true);
                  setSelectedCategory("ALL");
                  setCurrentPage(1);
                  setSearchFocused(true);
                }
              }}
              onKeyDown={e => {
                if (!showSearchSuggestions) {
                  if (e.key === "Enter" && search.trim()) {
                    scrollToSearchResults();
                  }
                  return;
                }

                if (e.key === "ArrowDown") {
                  e.preventDefault();
                  setActiveSearchSuggestion(index =>
                    index < searchSuggestions.length - 1 ? index + 1 : 0
                  );
                }

                if (e.key === "ArrowUp") {
                  e.preventDefault();
                  setActiveSearchSuggestion(index =>
                    index > 0 ? index - 1 : searchSuggestions.length - 1
                  );
                }

                if (e.key === "Enter") {
                  e.preventDefault();

                  if (activeSearchSuggestion >= 0) {
                    selectSearchSuggestion(
                      searchSuggestions[activeSearchSuggestion]
                    );
                  } else {
                    scrollToSearchResults();
                  }
                }

                if (e.key === "Escape") {
                  setSearchFocused(false);
                  setActiveSearchSuggestion(-1);
                }
              }}
              placeholder="Buscar productos, marcas o categorías..."
              aria-autocomplete="list"
              aria-controls="search-suggestions"
              aria-expanded={showSearchSuggestions}
            />

            {showSearchSuggestions && (
              <div
                id="search-suggestions"
                className="search-suggestions"
                role="listbox"
                aria-label="Sugerencias de búsqueda"
              >
                {searchSuggestions.map((product, index) => (
                  <button
                    key={product.id}
                    type="button"
                    className={`search-suggestion ${
                      activeSearchSuggestion === index ? "active" : ""
                    }`}
                    onMouseDown={e => e.preventDefault()}
                    onClick={() => selectSearchSuggestion(product)}
                    role="option"
                    aria-selected={activeSearchSuggestion === index}
                  >
                    <Search size={23} />
                    <span>{highlightSearchMatch(product.name)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>

        <div className="info-trigger-wrap">
          <button
            className="info-trigger"
            onClick={() => setInfoOpen(true)}
            type="button"
          >
            <Truck size={18} />
            <span>Despacho y medios de pago</span>
            <ChevronRight size={17} />
          </button>
        </div>

        <section className="content">
          {loading && <div className="state"><LoaderCircle className="spin" /> Cargando catálogo...</div>}
          {error && <div className="state error"><AlertCircle /> {error}</div>}

          {!loading && !error && (
            <>
              {showHomePromos && (
                <section className="promo-section">
                  <div className="section-head">
                    <h2>Ofertas y productos destacados</h2>
                  </div>

                {products.some(
                  p => p.family.toLowerCase() === "hielos"
                ) && (
                    <div className="ice-deal">
                      <img
                        src="/hielo-volumen-banner.png"
                        alt="Descuento por volumen en hielo"
                      />
                    </div>
                  )}

                <div className="promo-carousel">
                  {promoTotalPages > 1 && (
                    <button
                      type="button"
                      className="promo-arrow promo-arrow-left"
                      onClick={() => changePromoPage(-1)}
                      aria-label="Promociones anteriores"
                    >
                      ←
                    </button>
                  )}

                  <div className="product-grid promo-grid" key={promoPage}>
                    {visibleFeaturedProducts.map(p => (
                      <ProductCard key={p.id} product={p} onAdd={addToCart} onImageClick={(product) => { setImageProduct(product); setImageZoom(1); }} promo={false} />
                    ))}
                  </div>

                  {promoTotalPages > 1 && (
                    <button
                      type="button"
                      className="promo-arrow promo-arrow-right"
                      onClick={() => changePromoPage(1)}
                      aria-label="Siguientes promociones"
                    >
                      →
                    </button>
                  )}
                </div>

                {promoTotalPages > 1 && (
                  <div className="promo-dots" aria-label="Páginas de promociones">
                    {Array.from({ length: promoTotalPages }).map((_, index) => (
                      <button
                        key={index}
                        type="button"
                        className={index === promoPage ? "active" : ""}
                        onClick={() => setPromoPage(index)}
                        aria-label={`Ver promociones ${index + 1}`}
                      />
                    ))}
                  </div>
                )}
                </section>
              )}

              <section id="products-section" className="products-head">
                <div className="section-head">
                  <h2>{
                    searchCategory
                      ? categoryGroups.find(c => c.key === searchCategory)?.label
                      : selectedCategory === "ALL"
                        ? "Todos los productos"
                        : categoryGroups.find(c => c.key === selectedCategory)?.label
                  }</h2>
                  <span>{filteredProducts.length} productos</span>
                </div>

                {categoryForDisplay !== "ALL" && (
                  <div className="catalog-controls">
                    {availableFamilies.length > 0 && (
                      <div className="catalog-family-filter">
                        <label htmlFor="catalog-family-select">Familia</label>
                        <div className="catalog-family-select-wrap">
                          <select
                            id="catalog-family-select"
                            value={selectedFamily}
                            onChange={(event) => selectFamily(event.target.value)}
                          >
                            <option value="ALL">TODOS</option>
                            {availableFamilies.map((family) => (
                              <option key={family} value={family}>{family}</option>
                            ))}
                          </select>
                        </div>
                      </div>
                    )}

                    <div className="catalog-family-filter catalog-sort-filter">
                      <label htmlFor="catalog-sort-select">Ordenar por</label>
                      <div className="catalog-family-select-wrap">
                        <select
                          id="catalog-sort-select"
                          value={catalogSort}
                          onChange={(event) => {
                            setCatalogSort(event.target.value);
                            setSearchSelectedProductId(null);
                            setCurrentPage(1);
                          }}
                        >
                          <option value="DEFAULT">Seleccionar</option>
                          <option value="BEST_SELLERS">Más vendidos</option>
                          <option value="PRICE_ASC">Precio: menor a mayor</option>
                          <option value="PRICE_DESC">Precio: mayor a menor</option>
                          <option value="NAME">Nombre: A–Z</option>
                        </select>
                      </div>
                    </div>
                  </div>
                )}

                {totalPages > 1 && (
                  <div className="pagination minimal-pagination">
                    <button
                      disabled={currentPage === 1}
                      onClick={() => {
                        setCurrentPage(page => page - 1);
                        document.getElementById("products-section")?.scrollIntoView({
                          behavior: "smooth",
                          block: "start"
                        });
                      }}
                    >
                      ← Anterior
                    </button>

                    <span className="page-indicator">
                      Página <strong>{currentPage}</strong> de <strong>{totalPages}</strong>
                    </span>

                    <button
                      disabled={currentPage === totalPages}
                      onClick={() => {
                        setCurrentPage(page => page + 1);
                        document.getElementById("products-section")?.scrollIntoView({
                          behavior: "smooth",
                          block: "start"
                        });
                      }}
                    >
                      Siguiente →
                    </button>
                  </div>
                )}

                {filteredProducts.length === 0 ? (
                  <div className="no-results">No encontramos productos con esos filtros.</div>
                ) : (
                  <>
                    <div className="product-grid">
                      {paginatedProducts.map(p => (
                        <ProductCard key={p.id} product={p} onAdd={addToCart} onImageClick={(product) => { setImageProduct(product); setImageZoom(1); }} searchResult={p.id === searchSelectedProductId} />
                      ))}
                    </div>

                    {totalPages > 1 && (
                      <div className="pagination minimal-pagination">
                        <button
                          disabled={currentPage === 1}
                          onClick={() => {
                            setCurrentPage(page => page - 1);
                            document.getElementById("products-section")?.scrollIntoView({
                              behavior: "smooth",
                              block: "start"
                            });
                          }}
                        >
                          ← Anterior
                        </button>

                        <span className="page-indicator">
                          Página <strong>{currentPage}</strong> de <strong>{totalPages}</strong>
                        </span>

                        <button
                          disabled={currentPage === totalPages}
                          onClick={() => {
                            setCurrentPage(page => page + 1);
                            document.getElementById("products-section")?.scrollIntoView({
                              behavior: "smooth",
                              block: "start"
                            });
                          }}
                        >
                          Siguiente →
                        </button>
                      </div>
                    )}


                  </>
                )}
              </section>
            </>
          )}
        </section>
      </main>

      {imageProduct && (
        <div
          className="product-image-modal-backdrop"
          role="dialog"
          aria-modal="true"
          aria-label={`Imagen ampliada de ${imageProduct.name}`}
          onClick={() => {
            setImageProduct(null);
            setImageZoom(1);
          }}
        >
          <div
            className="product-image-modal"
            onClick={e => e.stopPropagation()}
          >
            <button
              type="button"
              className="product-image-modal-close"
              onClick={() => {
                setImageProduct(null);
                setImageZoom(1);
              }}
              aria-label="Cerrar imagen"
            >
              ×
            </button>

            <div className="product-image-modal-top">
              <div className="product-image-modal-title">
                <strong>{imageProduct.name}</strong>
                <span>{imageProduct.family || imageProduct.category}</span>
              </div>

              <div className="product-image-zoom-controls" aria-label="Controles de zoom">
                <button
                  type="button"
                  onClick={() => setImageZoom(z => Math.max(1, +(z - 0.25).toFixed(2)))}
                  disabled={imageZoom <= 1}
                  aria-label="Alejar"
                >
                  −
                </button>
                <span>{Math.round(imageZoom * 100)}%</span>
                <button
                  type="button"
                  onClick={() => setImageZoom(z => Math.min(2.5, +(z + 0.25).toFixed(2)))}
                  disabled={imageZoom >= 2.5}
                  aria-label="Acercar"
                >
                  +
                </button>
              </div>
            </div>

            <div
              className="product-image-modal-stage"
              onWheel={(e) => {
                e.preventDefault();
                setImageZoom(z => {
                  const next = z + (e.deltaY < 0 ? 0.15 : -0.15);
                  return Math.min(2.5, Math.max(1, +next.toFixed(2)));
                });
              }}
              onDoubleClick={() => setImageZoom(z => z > 1 ? 1 : 2)}
            >
              <img
                src={imageProduct.image}
                alt={imageProduct.name}
                className="product-image-modal-img"
                style={{ transform: `scale(${imageZoom})` }}
                draggable="false"
              />
            </div>

            <div className="product-image-modal-actions">
              <div className="product-image-modal-price">
                <span>{money(imageProduct.price)}</span>
                <small className={imageProduct.stock > 0 ? "stock-ok" : "stock-no"}>
                  {imageProduct.stock > 0 ? "• En stock" : "• Sin stock"}
                </small>
              </div>
              <button
                type="button"
                className="product-image-modal-add"
                disabled={imageProduct.stock <= 0}
                onClick={addFromImageModal}
              >
                <ShoppingCart size={17} />
                <span>{imageProduct.stock > 0 ? "Agregar al pedido" : "Sin stock"}</span>
              </button>
            </div>

            <div className="product-image-modal-footer">
              <span>Usa + / − o la rueda del mouse para ampliar</span>
              <span className="product-image-modal-hint-mobile">Toca + / − para ampliar</span>
            </div>
          </div>
        </div>
      )}

      {!cartOpen && !imageProduct && (
        <button
          className="cart-toggle"
          onClick={() => setCartOpen(true)}
          aria-label={`Ver pedido, ${cartUnits} productos`}
        >
          <ShoppingCart size={23} />
          <span>Ver pedido</span>
          <span className="cart-toggle-whatsapp" aria-label="WhatsApp">
            <svg viewBox="0 0 32 32" aria-hidden="true" focusable="false">
              <path d="M16 2.6C8.63 2.6 2.65 8.58 2.65 15.95c0 2.35.62 4.55 1.79 6.49L2.6 29.4l7.13-1.82a13.28 13.28 0 0 0 6.27 1.57h.01c7.36 0 13.34-5.98 13.34-13.34C29.35 8.58 23.37 2.6 16 2.6Z" fill="#25D366"/>
              <path d="M11.07 8.86c.3-.43.75-.64 1.26-.56l1.56.24c.43.07.77.35.9.77l.54 1.73c.11.36.02.75-.25 1.02l-.77.78c-.18.18-.22.45-.09.68.74 1.3 1.77 2.38 3.03 3.19.23.15.52.14.72-.04l.82-.74c.29-.26.71-.34 1.06-.2l1.63.66c.4.16.67.53.68.96l.04 1.54c.01.5-.25.96-.69 1.2-.58.32-1.28.48-1.98.4-1.76-.2-3.71-1.32-5.55-3.16-1.83-1.83-2.96-3.78-3.16-5.54-.08-.71.08-1.41.4-1.99l.85-.94Z" fill="#fff"/>
            </svg>
          </span>
          <b>{cartUnits}</b>
        </button>
      )}

      <aside ref={cartRef} className={`cart ${cartOpen ? "cart-open" : ""}`}>
        <div className="cart-box">
          <div className="cart-head">
            <ShoppingCart size={32} />
            <strong>Tu pedido</strong>
            <span>{cartUnits}</span>
            <button
              className="cart-close"
              onClick={() => setCartOpen(false)}
              aria-label="Cerrar pedido"
            >
              ×
            </button>
          </div>
          <div className="cart-items">
            {cart.length === 0 ? <p>Tu pedido aparecerá aquí.</p> : cart.map(item => (
              <div className="cart-row" key={item.id}>
                <div className="cart-thumb">{item.image ? <img className="cart-product-image" src={item.image} alt="" loading="lazy" onError={(e) => { e.currentTarget.style.display = "none"; }} /> : <Package size={22} />}</div>
                <div className="cart-info"><b>{item.name}</b><strong>{money(item.price)}</strong></div>
                <div className="qty">
                  <button onClick={() => changeQty(item.id, -1)} aria-label="Disminuir cantidad"><Minus size={14} /></button>
                  <QuantityInput
                    value={item.quantity}
                    max={Math.floor(item.stock)}
                    onChange={(value) => setQty(item.id, value)}
                  />
                  <button onClick={() => changeQty(item.id, 1)} aria-label="Aumentar cantidad"><Plus size={14} /></button>
                </div>
                <button className="remove" onClick={() => changeQty(item.id, -item.quantity)}><Trash2 size={18} /></button>
              </div>
            ))}
          </div>
          <div className="cart-total">
            <div>
              <span>Total ({cartUnits} productos)</span>
              {iceDiscount > 0 && (
                <small>
                  Descuento hielo: -{money(iceDiscount)}
                </small>
              )}
            </div>
            <b>{money(cartTotal)}</b>
          </div>
          <button className="whatsapp" onClick={sendWhatsApp}>
            <span className="order-whatsapp-icon" aria-hidden="true">
              <svg viewBox="0 0 32 32">
                <path
                  d="M16 2.6C8.63 2.6 2.65 8.58 2.65 15.95c0 2.35.62 4.55 1.79 6.49L2.6 29.4l7.13-1.82a13.28 13.28 0 0 0 6.27 1.57h.01c7.36 0 13.34-5.98 13.34-13.34C29.35 8.58 23.37 2.6 16 2.6Z"
                  fill="#25D366"
                />
                <path
                  d="M11.07 8.86c.3-.43.75-.64 1.26-.56l1.56.24c.43.07.77.35.9.77l.54 1.73c.11.36.02.75-.25 1.02l-.77.78c-.18.18-.22.45-.09.68.74 1.3 1.77 2.38 3.03 3.19.23.15.52.14.72-.04l.82-.74c.29-.26.71-.34 1.06-.2l1.63.66c.4.16.67.53.68.96l.04 1.54c.01.5-.25.96-.69 1.2-.58.32-1.28.48-1.98.4-1.76-.2-3.71-1.32-5.55-3.16-1.83-1.83-2.96-3.78-3.16-5.54-.08-.71.08-1.41.4-1.99l.85-.94Z"
                  fill="#fff"
                />
              </svg>
            </span>
            <span>Enviar pedido por WhatsApp</span>
          </button>
          <button className="continue-catalog" onClick={() => setCartOpen(false)}>
            <ChevronRight size={19} />
            <span>Seguir viendo el catálogo</span>
          </button>
          <button className="clear" onClick={() => setCart([])}><Trash2 size={20} /> Vaciar carrito</button>
        </div>
      </aside>

      {infoOpen && (
        <div className="info-modal-backdrop" onClick={() => setInfoOpen(false)}>
          <div
            className="info-modal info-modal-render"
            role="dialog"
            aria-modal="true"
            aria-labelledby="info-modal-title"
            onClick={e => e.stopPropagation()}
          >
            <img
              className="info-modal-render-image"
              src="/el-palmar-popup-inicio.png"
              alt="Información sobre despachos y medios de pago de Distribuidora El Palmar"
            />

            <button
              type="button"
              className="popup-hotspot popup-hotspot-close"
              onClick={() => setInfoOpen(false)}
              aria-label="Cerrar información"
            >
              <span aria-hidden="true">×</span>
            </button>

            <button
              type="button"
              className="popup-hotspot popup-hotspot-whatsapp"
              onClick={openWhatsAppContact}
              aria-label="Abrir WhatsApp para consultas"
            >
              <span className="sr-only">Consultas por WhatsApp</span>
            </button>

            <button
              type="button"
              className="popup-hotspot popup-hotspot-home"
              onClick={() => {
                setInfoOpen(false);
                goHome();
              }}
              aria-label="Ir al inicio"
            >
              <span className="sr-only">Ir al inicio</span>
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function CreditCardIcon() {
  return (
    <svg
      width="23"
      height="23"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <rect x="3" y="5" width="18" height="14" rx="2" />
      <path d="M3 10h18" />
      <path d="M7 15h3" />
    </svg>
  );
}

function QuantityInput({ value, max, onChange }) {
  const [draft, setDraft] = useState(String(value));

  useEffect(() => {
    setDraft(String(value));
  }, [value]);

  function handleChange(e) {
    const raw = e.target.value.replace(/\\D/g, "");
    setDraft(raw);

    if (raw !== "") {
      onChange(Number(raw));
    }
  }

  function handleBlur() {
    const number = Math.min(Math.max(1, Number(draft) || 1), max || 1);
    setDraft(String(number));
    onChange(number);
  }

  function handleKeyDown(e) {
    if (e.key === "Enter") {
      e.currentTarget.blur();
    }
  }

  return (
    <input
      className="qty-input"
      type="text"
      inputMode="numeric"
      value={draft}
      onChange={handleChange}
      onBlur={handleBlur}
      onKeyDown={handleKeyDown}
      aria-label="Cantidad"
    />
  );
}

function ProductCard({ product, onAdd, onImageClick, promo = false, searchResult = false }) {
  const available = product.stock > 0;
  return (
    <article className={`card ${searchResult ? "search-result-card" : ""}`} id={`product-${product.id}`}>
      {searchResult && (
        <div className="search-result-badge">
          <Search size={13} />
          <span>Resultado de búsqueda</span>
        </div>
      )}
      <div
        className={`product-placeholder ${product.image ? "product-image-clickable" : ""}`}
        onClick={() => product.image && onImageClick?.(product)}
        role={product.image ? "button" : undefined}
        tabIndex={product.image ? 0 : undefined}
        onKeyDown={(e) => {
          if (product.image && (e.key === "Enter" || e.key === " ")) {
            e.preventDefault();
            onImageClick?.(product);
          }
        }}
        aria-label={product.image ? `Ver imagen ampliada de ${product.name}` : undefined}
      >
        {product.image ? (
          <img
            className="product-image"
            src={product.image}
            alt={product.name}
            loading="lazy"
            onError={(e) => { e.currentTarget.style.display = "none"; }}
          />
        ) : (
          <div className="product-photo-placeholder" aria-label="Foto en proceso">
            <Camera size={30} strokeWidth={1.6} />
            <span>FOTO EN PROCESO</span>
            <small>Disponible próximamente</small>
          </div>
        )}
        {product.image && (
          <span className="product-image-zoom-badge" aria-hidden="true">⌕</span>
        )}
      </div>
      <div className="product-info">
        {promo && <em className="promo">PROMOCIÓN</em>}
        <b title={product.name}>{product.name}</b>
        <span>{product.family || product.category}</span>
        <strong>{money(product.price)}</strong>
        <em className={available ? "stock-ok" : "stock-no"}>{available ? "• En stock" : "• Sin stock"}</em>
        <button disabled={!available} onClick={() => onAdd(product)}><ShoppingCart size={13} /> Agregar</button>
      </div>
    </article>
  );
}

createRoot(document.getElementById("root")).render(<App />);
