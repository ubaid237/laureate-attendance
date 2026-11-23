# ==============================================================================
# 🛡️ CONFIDENTIAL: GHOST PROTOCOL / VIP REMIT CONFIGURATION
# (Yeh file website par kisi user ya visitor ko bilkul nazar NAHI aayegi)
# ==============================================================================
# Yahan aap har campus k VIP / Remit IDs set kr skte hain.
#
# Format:
# 1. "General": Jo IDs tamam campuses par lagani hon (e.g. Director, Owner)
# 2. "Campus Name": Specific campus k IDs (e.g. ["1", "5"])
#
# Note:
# - Campus ka naam exact match hona zaroori nahi, partial name b chalega.
#   Misal k tor pr: "Haris" likhenge toh "Haris Campus", "Haris Campus Multan" par khud match ho jaye ga.
# - IDs ko quotation marks mai list mai likhein, e.g. ["30", "66"]
# ==============================================================================

VIP_CONFIG = {
    # Sab campuses k liye (Agar koi general ID ho):
    "General": [],
    
    # Haris Campus k liye ID 30 aur 66:
    "Haris": ["30", "66"],
    
    # Mazeed campuses aap yahan add kr skte hain, jaise:
    # "Boys Campus": ["1", "5"],
    # "Cantt": ["7"],
}
