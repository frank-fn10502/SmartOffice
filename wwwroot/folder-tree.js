// SmartOffice Dashboard - Folder Tree Logic
(function() {
    'use strict';
    
    window.FolderTreeManager = {
        expandStates: {},
        selectedPath: '',
        
        // System folders to hide
        hiddenFolderNames: [
            'common views', 'finder', 'reminders', 'quick step',
            'conversation history', 'conversation action',
            'server failures', 'local failures', 'conflicts',
            'sync issues', 'rss', 'social network', 'people',
            'externalcontacts', 'yammer', 'RSS'
        ],
        
        isHiddenFolder: function(name) {
            var lowerName = (name || '').toLowerCase();
            return this.hiddenFolderNames.some(function(hidden) {
                return lowerName.includes(hidden);
            });
        },
        
        getFolderType: function(name) {
            var lowerName = (name || '').toLowerCase();
            
            if (lowerName === 'inbox' || lowerName === '收件匣' || lowerName === '收件箱') return 'inbox';
            if (lowerName === 'sent items' || lowerName === '寄件備份' || lowerName.includes('sent')) return 'sent';
            if (lowerName === 'drafts' || lowerName === '草稿') return 'drafts';
            if (lowerName === 'deleted items' || lowerName === '刪除的郵件' || lowerName.includes('deleted')) return 'deleted';
            if (lowerName === 'junk email' || lowerName === 'junk e-mail' || lowerName === '垃圾郵件') return 'junk';
            if (lowerName === 'archive' || lowerName === '封存') return 'archive';
            if (lowerName === 'outbox' || lowerName === '寄件匣') return 'outbox';
            
            return 'normal';
        },
        
        getFolderIcon: function(type) {
            var icons = {
                'inbox': '&#x1F4E5;',      // ??
                'sent': '&#x1F4E4;',       // ??
                'drafts': '&#x1F4DD;',     // ??
                'deleted': '&#x1F5D1;',    // ???
                'junk': '&#x1F6AB;',       // ??
                'archive': '&#x1F4E6;',    // ??
                'outbox': '&#x1F4EE;',     // ??
                'normal': '&#x1F4C1;'      // ??
            };
            return icons[type] || icons.normal;
        },
        
        createFolderElement: function(folder, level) {
            if (this.isHiddenFolder(folder.name)) {
                return null;
            }
            
            var type = this.getFolderType(folder.name);
            var visibleChildren = this.getVisibleChildren(folder);
            var hasChildren = visibleChildren.length > 0;
            var isExpanded = this.expandStates[folder.folderPath] || false;
            
            var container = document.createElement('div');
            container.className = 'folder-container';
            
            // Folder item
            var item = document.createElement('div');
            item.className = 'folder-item ' + type;
            item.setAttribute('data-path', folder.folderPath);
            
            // Chevron
            var chevron = document.createElement('span');
            chevron.className = 'folder-chevron' + (hasChildren ? (isExpanded ? ' expanded' : '') : ' empty');
            chevron.innerHTML = hasChildren ? '&#x25B6;' : ''; // ?
            
            var self = this;
            chevron.onclick = function(e) {
                e.stopPropagation();
                if (hasChildren) {
                    self.toggleFolder(folder.folderPath, container);
                }
            };
            
            // Icon
            var icon = document.createElement('span');
            icon.className = 'folder-icon';
            icon.innerHTML = this.getFolderIcon(type);
            
            // Name
            var nameSpan = document.createElement('span');
            nameSpan.className = 'folder-name';
            nameSpan.textContent = folder.name;
            
            // Count
            var countSpan = document.createElement('span');
            countSpan.className = 'folder-count';
            countSpan.textContent = '(' + folder.itemCount + ')';
            
            item.appendChild(chevron);
            item.appendChild(icon);
            item.appendChild(nameSpan);
            item.appendChild(countSpan);
            
            item.onclick = function(e) {
                if (e.target === chevron) return;
                self.selectFolder(folder.folderPath);
            };
            
            container.appendChild(item);
            
            // Children
            if (hasChildren) {
                var childrenDiv = document.createElement('div');
                childrenDiv.className = 'folder-children' + (isExpanded ? '' : ' collapsed');
                childrenDiv.setAttribute('data-path', folder.folderPath);
                
                visibleChildren.forEach(function(child) {
                    var childEl = self.createFolderElement(child, level + 1);
                    if (childEl) childrenDiv.appendChild(childEl);
                });
                
                container.appendChild(childrenDiv);
            }
            
            return container;
        },
        
        getVisibleChildren: function(folder) {
            if (!folder.subFolders || !folder.subFolders.length) return [];
            
            var self = this;
            return folder.subFolders.filter(function(child) {
                return !self.isHiddenFolder(child.name);
            });
        },
        
        toggleFolder: function(path, container) {
            this.expandStates[path] = !this.expandStates[path];
            var isExpanded = this.expandStates[path];
            
            var chevron = container.querySelector('.folder-chevron');
            var children = container.querySelector('.folder-children');
            
            if (chevron && children) {
                if (isExpanded) {
                    chevron.classList.add('expanded');
                    children.classList.remove('collapsed');
                } else {
                    chevron.classList.remove('expanded');
                    children.classList.add('collapsed');
                }
            }
        },
        
        selectFolder: function(path) {
            this.selectedPath = path;
            
            document.querySelectorAll('.folder-item.selected').forEach(function(el) {
                el.classList.remove('selected');
            });
            
            var item = document.querySelector('.folder-item[data-path="' + path.replace(/"/g, '\\"') + '"]');
            if (item) {
                item.classList.add('selected');
            }
        },
        
        renderTree: function(folders, container) {
            if (!folders || !folders.length) {
                container.innerHTML = '<span class="hint">No folders found.</span>';
                return;
            }
            
            container.innerHTML = '';
            
            var self = this;
            folders.forEach(function(rootFolder) {
                // Skip root, render children
                if (rootFolder.subFolders && rootFolder.subFolders.length > 0) {
                    rootFolder.subFolders.forEach(function(folder) {
                        var el = self.createFolderElement(folder, 0);
                        if (el) container.appendChild(el);
                    });
                } else if (!self.isHiddenFolder(rootFolder.name)) {
                    var el = self.createFolderElement(rootFolder, 0);
                    if (el) container.appendChild(el);
                }
            });
            
            // Auto-select Inbox
            if (!this.selectedPath) {
                var inboxItem = container.querySelector('.folder-item.inbox');
                if (inboxItem) {
                    var path = inboxItem.getAttribute('data-path');
                    if (path) this.selectFolder(path);
                }
            }
        }
    };
})();
