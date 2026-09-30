from django.contrib import admin
from .models import RootRadiusMaster

@admin.register(RootRadiusMaster)
class RootRadiusMasterAdmin(admin.ModelAdmin):
    list_display = ('planning_material', 'thickness', 'sec_wt', 'std', 'root_radius')
    list_filter = ('planning_material', 'std')
    search_fields = ('planning_material', 'std')