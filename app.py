import os
import sys
import math
import random
import traceback
from flask import Flask, jsonify, request, render_template, redirect, url_for

# Ensure dependencies can be imported
try:
    import pandas as pd
    import openpyxl
    import numpy as np
except ImportError as e:
    print(f"Missing dependency: {e}. Please run via run_app.py to resolve dependencies.")
    sys.exit(1)

app = Flask(__name__, template_folder='templates', static_folder='static')
app.config['UPLOAD_FOLDER'] = os.path.dirname(os.path.abspath(__file__))
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB limit

DEFAULT_EXCEL = os.path.join(app.config['UPLOAD_FOLDER'], 'ARCHIVOS.xlsx')
UPLOADED_EXCEL = os.path.join(app.config['UPLOAD_FOLDER'], 'uploaded_data.xlsx')

def get_active_filepath():
    if os.path.exists(UPLOADED_EXCEL):
        return UPLOADED_EXCEL
    return DEFAULT_EXCEL

def normalize_keyword(kw):
    if kw is None:
        return ""
    kw = str(kw).strip()
    if not kw:
        return ""
    
    # Standardize common terms and abbreviations
    lower_kw = kw.lower()
    if lower_kw in ['co2', 'lca', 'gpc', 'scm', 'rha', 'wsa', 'qd', 'sf', 'ggbs', 'fa']:
        return kw.upper()
    if lower_kw == 'co emissions':
        return 'CO2 Emissions'
    if lower_kw == 'co2 emissions':
        return 'CO2 Emissions'
    if lower_kw == 'carbon dioxide':
        return 'CO2'
    if lower_kw in ['supplementary cementitious material', 'supplementary cementitious materials', 'supplementary cementitious materials (scm)', 'scm']:
        return 'Supplementary Cementitious Materials'
    if lower_kw in ['rice husk ash (rha)', 'rice husk ash']:
        return 'Rice Husk Ash'
    if lower_kw in ['compressive strength', 'compressive strengths']:
        return 'Compressive Strength'
    if lower_kw in ['flexural strength', 'flexural strengths']:
        return 'Flexural Strength'
    
    # Standard Title Case for keywords
    words = kw.split()
    capitalized_words = []
    for w in words:
        w_lower = w.lower()
        if 'co2' in w_lower:
            w = w_lower.replace('co2', 'CO2').upper()
        elif w_lower in ["and", "or", "of", "in", "for", "with", "on", "at", "by", "from", "to", "as", "the", "a"]:
            w = w_lower
        else:
            w = w.capitalize()
        capitalized_words.append(w)
    
    res = " ".join(capitalized_words)
    if res:
        res = res[0].upper() + res[1:]
    return res

def normalize_author(author):
    if author is None:
        return ""
    author = str(author).strip()
    if not author:
        return ""
    
    # Format: Lastname, Firstname Middle -> Lastname, F.M.
    parts = author.split(",")
    if len(parts) == 2:
        lastname = parts[0].strip()
        firstname = parts[1].strip()
        initials_list = [n[0].upper() + "." for n in firstname.split() if n]
        initials = "".join(initials_list)
        return f"{lastname}, {initials}"
    return author

def parse_excel_data(filepath):
    """
    Parses and merges all relevant worksheets in the given Excel workbook.
    Looks for standard bibliometric columns.
    """
    if not os.path.exists(filepath):
        return []

    wb = openpyxl.load_workbook(filepath, read_only=True)
    all_articles = []
    
    # We will look for sheets containing rows, skipping empty sheets like Hoja1
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        if sheet.max_row <= 1:
            continue
            
        # Read header row
        header = None
        for r in sheet.iter_rows(max_row=1, values_only=True):
            header = [str(x).strip() if x is not None else "" for x in r]
            break
            
        if not header or not any(header):
            continue
            
        # Clean header names for case-insensitive matching
        header_lower = [h.lower() for h in header]
        
        def find_index(possible_names):
            for name in possible_names:
                if name.lower() in header_lower:
                    return header_lower.index(name.lower())
            # Fallback to search substrings
            for i, h in enumerate(header_lower):
                for name in possible_names:
                    if name.lower() in h:
                        return i
            return -1

        idx_key = find_index(['key', 'id'])
        idx_title = find_index(['title', 'titulo', 'document title'])
        idx_author = find_index(['author', 'authors', 'autor', 'autores'])
        idx_journal = find_index(['publication title', 'journal', 'source title', 'revista'])
        idx_tags = find_index(['automatic tags', 'manual tags', 'keywords', 'palabras clave', 'tags'])
        idx_year = find_index(['publication year', 'year', 'año'])
        idx_abstract = find_index(['abstract note', 'abstract', 'resumen'])

        # Iterate data rows
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # Check if row is empty or key is missing
            key_val = row[idx_key] if idx_key != -1 and idx_key < len(row) else None
            title_val = row[idx_title] if idx_title != -1 and idx_title < len(row) else None
            
            if not key_val and not title_val:
                continue # Skip empty row
                
            # Extract fields
            authors_raw = row[idx_author] if idx_author != -1 and idx_author < len(row) else ""
            journal_raw = row[idx_journal] if idx_journal != -1 and idx_journal < len(row) else ""
            tags_raw = row[idx_tags] if idx_tags != -1 and idx_tags < len(row) else ""
            year_raw = row[idx_year] if idx_year != -1 and idx_year < len(row) else ""
            abstract_raw = row[idx_abstract] if idx_abstract != -1 and idx_abstract < len(row) else ""
            
            # Format and normalize
            # Key
            key = str(key_val).strip() if key_val else f"KEY_{random.randint(10000, 99999)}"
            # Title
            title = str(title_val).strip() if title_val else "Sin Título"
            # Authors
            authors = []
            if authors_raw:
                authors = [normalize_author(a) for a in str(authors_raw).split(';') if a.strip()]
            # Journal
            journal = str(journal_raw).strip() if journal_raw else "Revista Desconocida"
            # Keywords
            keywords = []
            if tags_raw:
                # Support both comma and semicolon split
                split_char = ';' if ';' in str(tags_raw) else ','
                # Handle Zotero placeholder keyword
                raw_keywords = [k.strip() for k in str(tags_raw).split(split_char) if k.strip()]
                for k in raw_keywords:
                    # Ignore template/placeholder tags
                    if "type your keywords here" in k.lower() or "separated by semicolons" in k.lower():
                        continue
                    norm_k = normalize_keyword(k)
                    if norm_k:
                        keywords.append(norm_k)
            # Year
            try:
                year = int(year_raw) if year_raw else None
            except:
                year = None
            # Abstract
            abstract = str(abstract_raw).strip() if abstract_raw else ""
            
            all_articles.append({
                "key": key,
                "title": title,
                "authors": authors,
                "journal": journal,
                "keywords": keywords,
                "year": year,
                "abstract": abstract
            })
            
    # Deduplicate articles based on key or title
    seen = set()
    dedup_articles = []
    for art in all_articles:
        uniq_id = art["key"] if art["key"] else art["title"].lower()
        if uniq_id not in seen:
            seen.add(uniq_id)
            dedup_articles.append(art)
            
    return dedup_articles

def run_label_propagation(nodes, edges):
    """
    Simple Label Propagation algorithm to cluster nodes dynamically.
    Returns a dictionary mapping node_id to cluster index (0, 1, 2...).
    """
    if not nodes:
        return {}
        
    adj = {n['id']: {} for n in nodes}
    for e in edges:
        s, t, w = e['source'], e['target'], e['weight']
        if s in adj and t in adj:
            adj[s][t] = w
            adj[t][s] = w
            
    # Initialize labels
    labels = {n['id']: i for i, n in enumerate(nodes)}
    node_ids = [n['id'] for n in nodes]
    
    # 10 iterations max
    for _ in range(10):
        random.shuffle(node_ids)
        changed = False
        for u in node_ids:
            if not adj[u]:
                continue
            # Weight labels by edge weight
            label_weights = {}
            for v, w in adj[u].items():
                lbl = labels[v]
                label_weights[lbl] = label_weights.get(lbl, 0) + w
            if label_weights:
                best_label = max(label_weights.items(), key=lambda x: x[1])[0]
                if labels[u] != best_label:
                    labels[u] = best_label
                    changed = True
        if not changed:
            break
            
    # Normalize labels to 0, 1, 2, 3...
    unique_labels = sorted(list(set(labels.values())))
    label_map = {lbl: idx for idx, lbl in enumerate(unique_labels)}
    return {node_id: label_map[lbl] for node_id, lbl in labels.items()}

def generate_networks(articles):
    """
    Generates nodes and edges for Keywords, Authors, and Journals from parsed articles.
    """
    kw_freq = {}
    auth_freq = {}
    journ_freq = {}
    
    kw_pairs = {}
    auth_pairs = {}
    
    # Journal keywords association: to build links between journals
    # We will map: journal -> set of keywords used in its articles
    journal_keywords = {}
    
    for art in articles:
        # Keywords frequencies and pairs
        kws = list(set(art["keywords"]))  # Unique keywords in this article
        for k in kws:
            kw_freq[k] = kw_freq.get(k, 0) + 1
        for i in range(len(kws)):
            for j in range(i+1, len(kws)):
                k1, k2 = sorted([kws[i], kws[j]])
                kw_pairs[(k1, k2)] = kw_pairs.get((k1, k2), 0) + 1
                
        # Authors frequencies and pairs
        auths = list(set(art["authors"]))
        for a in auths:
            auth_freq[a] = auth_freq.get(a, 0) + 1
        for i in range(len(auths)):
            for j in range(i+1, len(auths)):
                a1, a2 = sorted([auths[i], auths[j]])
                auth_pairs[(a1, a2)] = auth_pairs.get((a1, a2), 0) + 1
                
        # Journals frequencies
        j = art["journal"]
        if j:
            journ_freq[j] = journ_freq.get(j, 0) + 1
            if j not in journal_keywords:
                journal_keywords[j] = set()
            for k in kws:
                journal_keywords[j].add(k)
                
    # 1. Keywords Network
    kw_nodes = [{"id": k, "count": f} for k, f in kw_freq.items()]
    kw_edges = [{"source": p[0], "target": p[1], "weight": w} for p, w in kw_pairs.items()]
    
    # 2. Authors Network
    auth_nodes = [{"id": a, "count": f} for a, f in auth_freq.items()]
    auth_edges = [{"source": p[0], "target": p[1], "weight": w} for p, w in auth_pairs.items()]
    
    # 3. Journals Network
    journ_nodes = [{"id": j, "count": f} for j, f in journ_freq.items()]
    
    # Journal connections based on shared keywords
    journ_edges = []
    j_list = list(journal_keywords.keys())
    for i in range(len(j_list)):
        for j in range(i+1, len(j_list)):
            j1, j2 = j_list[i], j_list[j]
            shared_kws = journal_keywords[j1].intersection(journal_keywords[j2])
            weight = len(shared_kws)
            if weight > 0:
                journ_edges.append({"source": j1, "target": j2, "weight": weight})
                
    # Run dynamic clustering for coloring on Keywords
    kw_clusters = run_label_propagation(kw_nodes, kw_edges)
    for n in kw_nodes:
        n['cluster'] = kw_clusters.get(n['id'], 0)
        
    # Run clustering on Authors
    auth_clusters = run_label_propagation(auth_nodes, auth_edges)
    for n in auth_nodes:
        n['cluster'] = auth_clusters.get(n['id'], 0)
        
    # Run clustering on Journals
    journ_clusters = run_label_propagation(journ_nodes, journ_edges)
    for n in journ_nodes:
        n['cluster'] = journ_clusters.get(n['id'], 0)

    # Compile cluster labels (most frequent element in each cluster)
    def compute_cluster_labels(nodes):
        c_groups = {}
        for n in nodes:
            c = n['cluster']
            if c not in c_groups:
                c_groups[c] = []
            c_groups[c].append(n)
        labels = {}
        for c, n_list in c_groups.items():
            # Sort by count descending
            n_list.sort(key=lambda x: x['count'], reverse=True)
            top_term = n_list[0]['id']
            labels[c] = f"{top_term} & asociados" if len(n_list) > 1 else top_term
        return labels

    return {
        "keywords": {
            "nodes": sorted(kw_nodes, key=lambda x: x['count'], reverse=True),
            "edges": kw_edges,
            "cluster_labels": compute_cluster_labels(kw_nodes)
        },
        "authors": {
            "nodes": sorted(auth_nodes, key=lambda x: x['count'], reverse=True),
            "edges": auth_edges,
            "cluster_labels": compute_cluster_labels(auth_nodes)
        },
        "journals": {
            "nodes": sorted(journ_nodes, key=lambda x: x['count'], reverse=True),
            "edges": journ_edges,
            "cluster_labels": compute_cluster_labels(journ_nodes)
        }
    }

# ============================================================
# FLASK ROUTING AND CONTROLLERS
# ============================================================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data')
def get_data():
    """
    Returns full network structure and counts.
    """
    try:
        filepath = get_active_filepath()
        if not os.path.exists(filepath):
            return jsonify({
                "success": False,
                "error": f"Archivo Excel no encontrado. Por favor, sube un archivo."
            })
            
        articles = parse_excel_data(filepath)
        networks = generate_networks(articles)
        
        # Calculate general stats
        all_kws = set()
        all_authors = set()
        all_journals = set()
        for art in articles:
            for k in art['keywords']:
                all_kws.add(k)
            for a in art['authors']:
                all_authors.add(a)
            if art['journal']:
                all_journals.add(art['journal'])
                
        stats = {
            "articles": len(articles),
            "keywords": len(all_kws),
            "authors": len(all_authors),
            "journals": len(all_journals),
            "source_file": os.path.basename(filepath)
        }
        
        # In addition, we return details about articles for information panels
        return jsonify({
            "success": True,
            "stats": stats,
            "networks": networks,
            "articles": articles
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({
            "success": False,
            "error": str(e)
        })

@app.route('/api/matrix')
def get_matrix():
    """
    Returns the correlation or similarity matrix.
    Parameters:
    - entity: keywords | authors | journals
    - metric: cooccurrence | cosine | pearson
    - top: integer (default 20)
    """
    try:
        entity = request.args.get('entity', 'keywords')
        metric = request.args.get('metric', 'cooccurrence')
        top = int(request.args.get('top', '20'))
        
        filepath = get_active_filepath()
        if not os.path.exists(filepath):
            return jsonify({"success": False, "error": "No Excel file loaded"})
            
        articles = parse_excel_data(filepath)
        
        # 1. Identify all occurrences and compute frequencies
        item_freq = {}
        paper_items = [] # list of lists, representing items present in each paper
        
        for art in articles:
            if entity == 'keywords':
                items = list(set(art['keywords']))
            elif entity == 'authors':
                items = list(set(art['authors']))
            else: # journals
                items = [art['journal']] if art['journal'] else []
                
            paper_items.append(items)
            for item in items:
                item_freq[item] = item_freq.get(item, 0) + 1
                
        # 2. Get top N items
        sorted_items = sorted(item_freq.items(), key=lambda x: x[1], reverse=True)
        top_items = [item for item, freq in sorted_items[:top]]
        top_n = len(top_items)
        
        if top_n == 0:
            return jsonify({
                "success": True,
                "labels": [],
                "matrix": []
            })
            
        # 3. Calculate full co-occurrence matrix for all items to enable profile-based correlation
        # To make it simple, we build an item map
        all_items_list = list(item_freq.keys())
        item_to_idx = {item: idx for idx, item in enumerate(all_items_list)}
        num_all = len(all_items_list)
        
        # Build co-occurrence matrix C
        C = np.zeros((num_all, num_all))
        for items in paper_items:
            indices = [item_to_idx[it] for it in items if it in item_to_idx]
            for idx in indices:
                C[idx, idx] += 1
            for i in range(len(indices)):
                for j in range(i+1, len(indices)):
                    idx1, idx2 = indices[i], indices[j]
                    C[idx1, idx2] += 1
                    C[idx2, idx1] += 1
                    
        # 4. Extract matrix for Top N
        # We will build a matrix of size top_n x top_n
        top_indices = [item_to_idx[item] for item in top_items]
        
        result_matrix = []
        if metric == 'cooccurrence':
            # Raw co-occurrence counts
            for i in top_indices:
                row = []
                for j in top_indices:
                    row.append(int(C[i, j]))
                result_matrix.append(row)
                
        elif metric == 'cosine':
            # Cosine similarity (Ochiai coefficient)
            for i in top_indices:
                row = []
                for j in top_indices:
                    freq_i = C[i, i]
                    freq_j = C[j, j]
                    if freq_i > 0 and freq_j > 0:
                        cos = C[i, j] / math.sqrt(freq_i * freq_j)
                        row.append(round(cos, 4))
                    else:
                        row.append(0.0)
                result_matrix.append(row)
                
        elif metric == 'pearson':
            # Pearson correlation of co-occurrence profiles (rows of C)
            # The correlation is computed using all columns in the co-occurrence profile
            # to capture global relationships!
            for i in top_indices:
                row = []
                profile_i = C[i, :]
                for j in top_indices:
                    profile_j = C[j, :]
                    
                    # Pearson correlation formula
                    mean_i = np.mean(profile_i)
                    mean_j = np.mean(profile_j)
                    
                    dev_i = profile_i - mean_i
                    dev_j = profile_j - mean_j
                    
                    sq_sum_i = np.sum(dev_i ** 2)
                    sq_sum_j = np.sum(dev_j ** 2)
                    
                    if sq_sum_i > 0 and sq_sum_j > 0:
                        corr = np.sum(dev_i * dev_j) / math.sqrt(sq_sum_i * sq_sum_j)
                        row.append(round(corr, 4))
                    else:
                        row.append(0.0)
                result_matrix.append(row)
                
        return jsonify({
            "success": True,
            "labels": top_items,
            "matrix": result_matrix,
            "entity": entity,
            "metric": metric
        })
        
    except Exception as e:
        traceback.print_exc()
        return jsonify({
            "success": False,
            "error": str(e)
        })

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """
    Endpoint for uploading a new Excel file.
    """
    if 'file' not in request.files:
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        return redirect(url_for('index'))
        
    if file and file.filename.endswith(('.xlsx', '.xls')):
        # Save file to uploaded path
        file.save(UPLOADED_EXCEL)
        
    return redirect(url_for('index'))

@app.route('/api/reset', methods=['POST'])
def reset_file():
    """
    Resets active Excel file to default ARCHIVOS.xlsx.
    """
    if os.path.exists(UPLOADED_EXCEL):
        os.remove(UPLOADED_EXCEL)
    return jsonify({"success": True})

if __name__ == '__main__':
    # Running directly (for development or simple deployments)
    port = int(os.environ.get('PORT', 5000))
    host = '0.0.0.0'
    app.run(debug=False, host=host, port=port)
